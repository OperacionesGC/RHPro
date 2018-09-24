Attribute VB_Name = "MdlFactoresTotalizadores"
Option Explicit

Public Sub SumaFactores(Nro_Analisis As Long, Filtrar As Boolean)
' ---------------------------------------------------------------------------------------------
' Descripcion: Analisis para tipo de factor 8 (Suma de Factores)
' Autor      : FGZ
' Fecha      : 01/12/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
'Variables locales
Dim cant_flt As Long
Dim Desde As Date
Dim Hasta As Date
Dim horas As Single
Dim NroCab As Long
Dim perpago_desde As Long
Dim perpago_hasta As Long

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
Dim Aux_cubvalor1 As Single
Dim Aux_cubvalor2 As Single


'Variables para los first y last
Dim PrimerFactOri As Boolean
Dim TipoEstr As Long
Dim EstrAct As Long
Dim FactOri As Long
Dim MiConcepto As String

Dim estr_liqNex As String
Dim cod_cptoNex As String

' Fechas desde y hasta a analizar por el proceso entero
Dim Inicio_Per_Analizado As Date
Dim Fin_Per_Analizado As Date

' Fechas parciales que se estan analizando
Dim Dia_Inicio_Per_Analizado As Date
Dim Dia_Fin_Per_Analizado As Date

' Auxiliares para el manejo de ls ciclos
Dim MesActual As Integer
Dim MesInicio As Integer
Dim MesFin As Integer
Dim AnioInicio As Integer
Dim AnioFin As Integer
Dim AnioActual As Integer
Dim AuxDia As Integer
Dim Ok As Boolean

Dim UltimoDiaMes As Integer
Dim AuxRangoDesde As Date
Dim AuxRangoHasta As Date

'FGZ - 11/09/2003
Dim NombreBD As String

Dim rs As New ADODB.Recordset
Dim rs_AnrCubo As New ADODB.Recordset

'Código -------------------------------------------------------------------
'Abro la conexion para nexus
'OpenConnection strConexionNexus, objConnNexus

On Error GoTo CE:

' --- fgz 07/07/2003---------
'Obtengo la cabecera y el filtro
Call ObtenerCabecerayFiltro(Nro_Analisis, rsAnrCab, Filtrar, rsFiltro, cant_flt, Ok)

If Not Ok Then
    Exit Sub
End If
CantFiltro = 0
If Not rsFiltro.EOF Then
    'CantFiltro = rsFiltro.RecordCount
    CantFiltro = cant_flt
Else
    CantFiltro = 1
End If


'Estoy reprocesando
Call PurgarCubo(Nro_Analisis, 8)
' ---------------------------

'Comienzo el procesamiento
'StrSql = "SELECT * FROM anrcab_fact" & _
'    " INNER JOIN anrfact_ori ON anrfact_ori.facnro = anrcab_fact.facnro" & _
'    " AND anrfact_ori.tipfacnro = 8" & _
'    " INNER JOIN anrfactor ON anrfactor.facnro = anrcab_fact.facnro" & _
'    " WHERE anrcabnro = " & rsAnrCab!anrcabnro & _
'    " ORDER BY anrfact_ori.facnro"
StrSql = "SELECT anrcab_fact.facnro, anrcab_fact.anrcabnro, anrfactor.tipfacnro, anrfactor.facpropor, "
StrSql = StrSql & " anrfact_tot.facnrotot, anrfactor.facpresup, anrfactor.facpresupmonto, anrfactor.facopsuma, anrfactor.facopfijo "
StrSql = StrSql & " FROM anrcab_fact"
StrSql = StrSql & " INNER JOIN anrfact_tot ON anrfact_tot.facnro = anrcab_fact.facnro"
StrSql = StrSql & " INNER JOIN anrfactor ON anrfact_tot.facnrotot = anrfactor.facnro"
StrSql = StrSql & " WHERE anrcab_fact.anrcabnro = " & Nro_Analisis
StrSql = StrSql & " AND anrfactor.tipfacnro = 8"
StrSql = StrSql & " ORDER BY anrcab_fact.nrocolum, anrcab_fact.facnro"
OpenRecordset StrSql, rsFactor
    
CantFactor = 0
If Not rsFactor.EOF Then
    CantFactor = rsFactor.RecordCount
Else
    CantFactor = 1
End If

'obtengo el conjunto de legajos a procesar
Inicio_Per_Analizado = rsAnrCab!anrcabfecdesde
Fin_Per_Analizado = rsAnrCab!anrcabfechasta
Call ObtenerLegajos(2, Filtrar, rsAnrCab!anrcabnro, rsFiltro, rsAnrCab!anrcabfecdesde, rsAnrCab!anrcabfechasta)

'Seteo el incremento del progreso
'Progreso = 95
Progreso = SumPorcTiempo
If (CantFactor * rsFiltro.RecordCount * CantFiltro) <> 0 Then
    IncPorc = PorcTiempo / (CantFactor * rsFiltro.RecordCount)
Else
    IncPorc = PorcTiempo
End If
                    
Do While Not rsFactor.EOF
'    'Busco si es factor totalizador
'    StrSql = "SELECT * FROM anrfact_tot" & _
'            " WHERE facnro = " & FactOri
'    OpenRecordset StrSql, rsFactorTotalizador
'    If Not rsFactorTotalizador.EOF Then
'        Totaliza = True
'        ' codigo de factor con el cual se inserta en el cubo
'        FactorTotalizador = rsFactorTotalizador!facnrotot
'    Else
'        Totaliza = False
'        FactorTotalizador = 0
'    End If
    
    'Primer legajo del filtro
    rsFiltro.MoveFirst
    
    Do While Not rsFiltro.EOF
        tercero = rsFiltro!Ternro
        If (cant_flt > 0) And (rsFiltro!cant_te < cant_flt) Then
            GoTo siguientelegajo
        End If
        'Cuando hay establecido un filtro, se debe verificar que el empleado verifique
        'todos los filtros en el intervalo de tiempo analizado. El control se hace de
        'esta forma, para considerar en forma correcta los casos en donde existe más de
        'un registro en HIS_ESTRUCTURA, para el mismo empleado y para el mismo tipo de
        'estructura, que satisfacen el intervalo de tiempo.
    
    
    
    
        ' Recorre para el analisis las tablas segun los factores configurados
        StrSql = "SELECT DISTINCT ternro, facnro, anrcubvalor1, anrcubvalor2, anrrangnro FROM anrcubo"
        StrSql = StrSql & " WHERE anrcabnro = " & Nro_Analisis
        StrSql = StrSql & " AND facnro = " & rsFactor!facnro
'        StrSql = StrSql & " AND tenro = " & rsEstructura!tenro
'        StrSql = StrSql & " AND estrnro = " & rsEstructura!estrnro
        StrSql = StrSql & " AND ternro = " & tercero
'        StrSql = StrSql & " AND anrrangnro = " & rsRango!anrrangnro
        OpenRecordset StrSql, rs_AnrCubo
        
        Do While Not rs_AnrCubo.EOF
                    
            StrSql = "SELECT * FROM anrrangofec"
            StrSql = StrSql & " WHERE anrrangofec.anrcabnro = " & rsAnrCab!anrcabnro
            StrSql = StrSql & " AND anrrangorepro = -1"
            StrSql = StrSql & " AND anrrangnro = " & rs_AnrCubo!anrrangnro
'            StrSql = StrSql & " AND anrrangofec.anrrangfecdesde <= " & ConvFecha(rs_detliq!profecini)
'            StrSql = StrSql & " AND anrrangofec.anrrangfechasta >= " & ConvFecha(rs_detliq!profecfin)
            OpenRecordset StrSql, rsRango
            
            Do While Not rsRango.EOF
                Fin_Per_Analizado = rsRango!anrrangfechasta
                Inicio_Per_Analizado = rsRango!anrrangfecdesde
                    
                ' FGZ 10/07/2003--------------------------
                Call ObtenerEstructuras(Filtrar, tercero, Inicio_Per_Analizado, Fin_Per_Analizado, rsEstructura)
                If Not rsEstructura.EOF Then
                    TipoEstr = rsEstructura!tenro
                    EstrAct = rsEstructura!estrnro
                End If
            
                Do While Not rsEstructura.EOF
                    If PrimerFactOri Then
                        cantdiasper = DateDiff("d", Inicio_Per_Analizado, Fin_Per_Analizado) + 1
                        monto_total = 0
                        cant_total = 0
                        cant_saldo = 0
                        PrimerFactOri = False
                    End If
                                           
                    'Acumulo por Factor
                    monto_total = monto_total + rs_AnrCubo!anrcubvalor1
                    cant_total = cant_total + rs_AnrCubo!anrcubvalor2
                    
                    'Calculo los dias de rango entre las fechas del rango y el his_estruct para proporcionar
                    If rsFactor!facpropor = -1 Then
                        If rsEstructura!htetdesde < Inicio_Per_Analizado Then
                            If rsEstructura!htethasta < Fin_Per_Analizado And (Not IsNull(rsEstructura!htethasta)) Then
                                cantdiasran = DateDiff("d", Inicio_Per_Analizado, rsEstructura!htethasta) + 1
                            Else
                                cantdiasran = DateDiff("d", Inicio_Per_Analizado, Fin_Per_Analizado) + 1
                            End If
                        Else
                            If (rsEstructura!htethasta < Fin_Per_Analizado) And (Not IsNull(rsEstructura!htethasta)) Then
                                cantdiasran = DateDiff("d", rsEstructura!htetdesde, rsEstructura!htethasta) + 1
                            Else
                                cantdiasran = DateDiff("d", rsEstructura!htetdesde, Fin_Per_Analizado) + 1
                            End If
                        End If
                                
                        'Porcentaje segun la cant. de dias en la his_estrutura
                        porcentaje = cantdiasran * 100 / cantdiasper
                                
                        If Last_OF_Factor() Or Last_OF_estrnro() Then
                            cubvalor1 = monto_total * porcentaje / 100
                            cubvalor2 = cant_total * porcentaje / 100
                                    
                            StrSql = "SELECT * FROM anrcubo"
                            StrSql = StrSql & " WHERE anrcabnro = " & Nro_Analisis
                            StrSql = StrSql & " AND facnro = " & rsFactor!facnrotot
                            StrSql = StrSql & " AND tenro = " & rsEstructura!tenro
                            StrSql = StrSql & " AND estrnro = " & rsEstructura!estrnro
                            StrSql = StrSql & " AND ternro = " & tercero
                            StrSql = StrSql & " AND anrrangnro = " & rsRango!anrrangnro
                            OpenRecordset StrSql, rs
        
                            'Si el cubo no existe lo creo
                            If rs.EOF Then
                                'Creo el cubo
                                StrSql = "INSERT INTO anrcubo(anrcabnro,anrcubmanual"
                                StrSql = StrSql & ",anrrangnro,estrnro,facnro,tenro,Ternro,tipnro"
                                StrSql = StrSql & ",anrcubvalor1,anrcubvalor2"
                                '---------------------------
                                ' FAF 14-02-2005
                                If CInt(rsFactor!facpresup) = -1 Then
                                    StrSql = StrSql & ",anrcubvalor3,anrcubvalor4"
                                End If
                                '---------------------------
                                StrSql = StrSql & ") VALUES ("
                                StrSql = StrSql & Nro_Analisis & ",0," & rsRango!anrrangnro & ","
                                StrSql = StrSql & rsEstructura!estrnro & "," & rsFactor!facnrotot & ","
                                StrSql = StrSql & rsEstructura!tenro & "," & tercero & ",1"
                            End If
                            monto_saldo = (monto_total - cubvalor1 - monto_saldo)
                            cant_saldo = (cant_total - cubvalor2 - cant_saldo)
                            monto_total = 0
                            cant_total = 0
                                    
                            'Para que no quede saldo cuando proporciona
                            If monto_saldo <= 1 And monto_saldo > 0 Then
                                cubvalor1 = cubvalor1 + monto_saldo
                            End If
                            If cant_saldo <= 1 And cant_saldo > 0 Then
                                'cubvalor2 = cubvalor2 + cant_saldo
                            End If
                                       
                            'Si existe el cubo entonces actualizo
                            If Not rs.EOF Then
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
                                    " AND facnro = " & rsFactor!facnrotot & _
                                    " AND tenro = " & rsEstructura!tenro & _
                                    " AND estrnro = " & rsEstructura!estrnro & _
                                    " AND ternro = " & tercero & _
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
                        'Si no proporciona tomo al 100% y la ultima his_estruc del rango
                        porcentaje = 100
                        'If Last_OF_Factor() Or Last_OF_estrnro() Then
                        If Ultimo(rs_AnrCubo) Or Last_OF_estrnro() Then
                            If Not Last_OF_tenro() Then
                                monto_total = 0
                                cant_total = 0
                            Else
                                'Busco la ultima his_estr dentro del rango
                                StrSql = "SELECT * FROM his_estructura " & _
                                    " WHERE his_estructura.ternro = " & tercero & _
                                    " AND his_estructura.tenro = " & rsEstructura!tenro & _
                                    " AND his_estructura.htetdesde <= " & ConvFecha(Fin_Per_Analizado) & _
                                    " AND (his_estructura.htethasta >= " & ConvFecha(Inicio_Per_Analizado) & _
                                    " OR his_estructura.htethasta IS NULL) "
                                OpenRecordset StrSql, objRs
                                objRs.MoveLast
                                
                                If Not objRs.EOF Then
                                    StrSql = "SELECT * FROM anrcubo" & _
                                        " WHERE anrcabnro = " & rsAnrCab!anrcabnro & _
                                        " AND facnro = " & rsFactor!facnrotot & _
                                        " AND tenro = " & objRs!tenro & _
                                        " AND estrnro = " & objRs!estrnro & _
                                        " AND ternro = " & tercero & _
                                        " AND anrrangnro = " & rsRango!anrrangnro
                                    OpenRecordset StrSql, rs
                                    
                                    cubvalor1 = monto_total * porcentaje / 100
                                    cubvalor2 = cant_total * porcentaje / 100
                                    
                                    If rs.EOF Then
                                        'Creo el cubo
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
                                            objRs!estrnro & "," & rsFactor!facnrotot & "," & _
                                            objRs!tenro & "," & tercero & ",1" & _
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
                                            Aux_cubvalor1 = cubvalor1
                                            Aux_cubvalor2 = cubvalor2
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
                                            " AND facnro = " & rsFactor!facnrotot & _
                                            " AND tenro = " & objRs!tenro & _
                                            " AND estrnro = " & objRs!estrnro & _
                                            " AND ternro = " & tercero & _
                                            " AND anrrangnro = " & rsRango!anrrangnro
                                            Aux_cubvalor1 = rs!anrcubvalor1 + cubvalor1
                                            Aux_cubvalor2 = rs!anrcubvalor2 + cubvalor2
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
            
                rsRango.MoveNext
            Loop
                
            ' Actualizo el progreso
            Progreso = Progreso + IncPorc
            StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & " WHERE bpronro = " & NroProceso
            objconnProgreso.Execute StrSql, , adExecuteNoRecords
                
            rs_AnrCubo.MoveNext
        Loop
        
siguientelegajo:
        rsFiltro.MoveNext
    Loop
    rsFactor.MoveNext
Loop
SumPorcTiempo = CDbl(SumPorcTiempo) + CDbl(PorcTiempo)
Exit Sub
CE:
    HuboErrorTipo = True
    HuboError = True
    Flog.writeline Espacios(Tabulador * 1) & "Error " & Err.Description
End Sub



Public Sub ProductoFactores(Nro_Analisis As Long, Filtrar As Boolean)
' ---------------------------------------------------------------------------------------------
' Descripcion: Analisis para tipo de factor 10 (Multiplicacion de Factores)
' Autor      : FGZ
' Fecha      : 01/12/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
'Variables locales
Dim cant_flt As Long
Dim Desde As Date
Dim Hasta As Date
Dim horas As Single
Dim NroCab As Long
Dim perpago_desde As Long
Dim perpago_hasta As Long

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
Dim Aux_cubvalor1 As Single
Dim Aux_cubvalor2 As Single


'Variables para los first y last
Dim PrimerFactOri As Boolean
Dim TipoEstr As Long
Dim EstrAct As Long
Dim FactOri As Long
Dim MiConcepto As String

Dim estr_liqNex As String
Dim cod_cptoNex As String

' Fechas desde y hasta a analizar por el proceso entero
Dim Inicio_Per_Analizado As Date
Dim Fin_Per_Analizado As Date

' Fechas parciales que se estan analizando
Dim Dia_Inicio_Per_Analizado As Date
Dim Dia_Fin_Per_Analizado As Date

' Auxiliares para el manejo de ls ciclos
Dim MesActual As Integer
Dim MesInicio As Integer
Dim MesFin As Integer
Dim AnioInicio As Integer
Dim AnioFin As Integer
Dim AnioActual As Integer
Dim AuxDia As Integer
Dim Ok As Boolean

Dim UltimoDiaMes As Integer
Dim AuxRangoDesde As Date
Dim AuxRangoHasta As Date

'FGZ - 11/09/2003
Dim NombreBD As String

Dim rs As New ADODB.Recordset
Dim rs_AnrCubo As New ADODB.Recordset

'Código -------------------------------------------------------------------
'Abro la conexion para nexus
'OpenConnection strConexionNexus, objConnNexus

On Error GoTo CE:

' --- fgz 07/07/2003---------
'Obtengo la cabecera y el filtro
Call ObtenerCabecerayFiltro(Nro_Analisis, rsAnrCab, Filtrar, rsFiltro, cant_flt, Ok)

If Not Ok Then
    Exit Sub
End If
CantFiltro = 0
If Not rsFiltro.EOF Then
    'CantFiltro = rsFiltro.RecordCount
    CantFiltro = cant_flt
Else
    CantFiltro = 1
End If


'Estoy reprocesando
Call PurgarCubo(Nro_Analisis, 10)
' ---------------------------

'Comienzo el procesamiento
'StrSql = "SELECT * FROM anrcab_fact" & _
'    " INNER JOIN anrfact_ori ON anrfact_ori.facnro = anrcab_fact.facnro" & _
'    " AND anrfact_ori.tipfacnro = 8" & _
'    " INNER JOIN anrfactor ON anrfactor.facnro = anrcab_fact.facnro" & _
'    " WHERE anrcabnro = " & rsAnrCab!anrcabnro & _
'    " ORDER BY anrfact_ori.facnro"
StrSql = "SELECT anrcab_fact.facnro, anrcab_fact.anrcabnro, anrfactor.tipfacnro, anrfactor.facpropor, "
StrSql = StrSql & " anrfact_tot.facnrotot, anrfactor.facpresup, anrfactor.facpresupmonto, anrfactor.facopsuma, anrfactor.facopfijo "
StrSql = StrSql & " FROM anrcab_fact"
StrSql = StrSql & " INNER JOIN anrfact_tot ON anrfact_tot.facnro = anrcab_fact.facnro"
StrSql = StrSql & " INNER JOIN anrfactor ON anrfact_tot.facnrotot = anrfactor.facnro"
StrSql = StrSql & " WHERE anrcab_fact.anrcabnro = " & Nro_Analisis
StrSql = StrSql & " AND anrfactor.tipfacnro = 10"
StrSql = StrSql & " ORDER BY anrcab_fact.nrocolum, anrcab_fact.facnro"
OpenRecordset StrSql, rsFactor
    
CantFactor = 0
If Not rsFactor.EOF Then
    CantFactor = rsFactor.RecordCount
Else
    CantFactor = 1
End If

'obtengo el conjunto de legajos a procesar
Inicio_Per_Analizado = rsAnrCab!anrcabfecdesde
Fin_Per_Analizado = rsAnrCab!anrcabfechasta
Call ObtenerLegajos(2, Filtrar, rsAnrCab!anrcabnro, rsFiltro, rsAnrCab!anrcabfecdesde, rsAnrCab!anrcabfechasta)

'Seteo el incremento del progreso
'Progreso = 95
Progreso = SumPorcTiempo
If (CantFactor * rsFiltro.RecordCount * CantFiltro) <> 0 Then
    IncPorc = PorcTiempo / (CantFactor * rsFiltro.RecordCount)
Else
    IncPorc = PorcTiempo
End If
                    
Do While Not rsFactor.EOF
'    'Busco si es factor totalizador
'    StrSql = "SELECT * FROM anrfact_tot" & _
'            " WHERE facnro = " & FactOri
'    OpenRecordset StrSql, rsFactorTotalizador
'    If Not rsFactorTotalizador.EOF Then
'        Totaliza = True
'        ' codigo de factor con el cual se inserta en el cubo
'        FactorTotalizador = rsFactorTotalizador!facnrotot
'    Else
'        Totaliza = False
'        FactorTotalizador = 0
'    End If
    
    'Primer legajo del filtro
    rsFiltro.MoveFirst
    
    Do While Not rsFiltro.EOF
        tercero = rsFiltro!Ternro
        If (cant_flt > 0) And (rsFiltro!cant_te < cant_flt) Then
            GoTo siguientelegajo
        End If
        'Cuando hay establecido un filtro, se debe verificar que el empleado verifique
        'todos los filtros en el intervalo de tiempo analizado. El control se hace de
        'esta forma, para considerar en forma correcta los casos en donde existe más de
        'un registro en HIS_ESTRUCTURA, para el mismo empleado y para el mismo tipo de
        'estructura, que satisfacen el intervalo de tiempo.
    
    
    
    
        ' Recorre para el analisis las tablas segun los factores configurados
        StrSql = "SELECT DISTINCT ternro, facnro, anrcubvalor1, anrcubvalor2, anrrangnro FROM anrcubo"
        StrSql = StrSql & " WHERE anrcabnro = " & Nro_Analisis
        StrSql = StrSql & " AND facnro = " & rsFactor!facnro
'        StrSql = StrSql & " AND tenro = " & rsEstructura!tenro
'        StrSql = StrSql & " AND estrnro = " & rsEstructura!estrnro
        StrSql = StrSql & " AND ternro = " & tercero
'        StrSql = StrSql & " AND anrrangnro = " & rsRango!anrrangnro
        OpenRecordset StrSql, rs_AnrCubo
        
        Do While Not rs_AnrCubo.EOF
                    
            StrSql = "SELECT * FROM anrrangofec"
            StrSql = StrSql & " WHERE anrrangofec.anrcabnro = " & rsAnrCab!anrcabnro
            StrSql = StrSql & " AND anrrangorepro = -1"
            StrSql = StrSql & " AND anrrangnro = " & rs_AnrCubo!anrrangnro
'            StrSql = StrSql & " AND anrrangofec.anrrangfecdesde <= " & ConvFecha(rs_detliq!profecini)
'            StrSql = StrSql & " AND anrrangofec.anrrangfechasta >= " & ConvFecha(rs_detliq!profecfin)
            OpenRecordset StrSql, rsRango
            
            Do While Not rsRango.EOF
                Fin_Per_Analizado = rsRango!anrrangfechasta
                Inicio_Per_Analizado = rsRango!anrrangfecdesde
                    
                ' FGZ 10/07/2003--------------------------
                Call ObtenerEstructuras(Filtrar, tercero, Inicio_Per_Analizado, Fin_Per_Analizado, rsEstructura)
                If Not rsEstructura.EOF Then
                    TipoEstr = rsEstructura!tenro
                    EstrAct = rsEstructura!estrnro
                End If
            
                Do While Not rsEstructura.EOF
                    If PrimerFactOri Then
                        cantdiasper = DateDiff("d", Inicio_Per_Analizado, Fin_Per_Analizado) + 1
                        monto_total = 0
                        cant_total = 0
                        cant_saldo = 0
                        PrimerFactOri = False
                    End If
                                           
                    'Acumulo por Factor
                    monto_total = monto_total + rs_AnrCubo!anrcubvalor1
                    cant_total = cant_total + rs_AnrCubo!anrcubvalor2
                    
                    'Calculo los dias de rango entre las fechas del rango y el his_estruct para proporcionar
                    If rsFactor!facpropor = -1 Then
                        If rsEstructura!htetdesde < Inicio_Per_Analizado Then
                            If rsEstructura!htethasta < Fin_Per_Analizado And (Not IsNull(rsEstructura!htethasta)) Then
                                cantdiasran = DateDiff("d", Inicio_Per_Analizado, rsEstructura!htethasta) + 1
                            Else
                                cantdiasran = DateDiff("d", Inicio_Per_Analizado, Fin_Per_Analizado) + 1
                            End If
                        Else
                            If (rsEstructura!htethasta < Fin_Per_Analizado) And (Not IsNull(rsEstructura!htethasta)) Then
                                cantdiasran = DateDiff("d", rsEstructura!htetdesde, rsEstructura!htethasta) + 1
                            Else
                                cantdiasran = DateDiff("d", rsEstructura!htetdesde, Fin_Per_Analizado) + 1
                            End If
                        End If
                                
                        'Porcentaje segun la cant. de dias en la his_estrutura
                        porcentaje = cantdiasran * 100 / cantdiasper
                                
                        If Last_OF_Factor() Or Last_OF_estrnro() Then
                            cubvalor1 = monto_total * porcentaje / 100
                            cubvalor2 = cant_total * porcentaje / 100
                                    
                            StrSql = "SELECT * FROM anrcubo"
                            StrSql = StrSql & " WHERE anrcabnro = " & Nro_Analisis
                            StrSql = StrSql & " AND facnro = " & rsFactor!facnrotot
                            StrSql = StrSql & " AND tenro = " & rsEstructura!tenro
                            StrSql = StrSql & " AND estrnro = " & rsEstructura!estrnro
                            StrSql = StrSql & " AND ternro = " & tercero
                            StrSql = StrSql & " AND anrrangnro = " & rsRango!anrrangnro
                            OpenRecordset StrSql, rs
        
                            'Si el cubo no existe lo creo
                            If rs.EOF Then
                                'Creo el cubo
                                StrSql = "INSERT INTO anrcubo(anrcabnro,anrcubmanual"
                                StrSql = StrSql & ",anrrangnro,estrnro,facnro,tenro,Ternro,tipnro"
                                StrSql = StrSql & ",anrcubvalor1,anrcubvalor2"
                                '---------------------------
                                ' FAF 14-02-2005
                                If CInt(rsFactor!facpresup) = -1 Then
                                    StrSql = StrSql & ",anrcubvalor3,anrcubvalor4"
                                End If
                                '---------------------------
                                StrSql = StrSql & ") VALUES ("
                                StrSql = StrSql & Nro_Analisis & ",0," & rsRango!anrrangnro & ","
                                StrSql = StrSql & rsEstructura!estrnro & "," & rsFactor!facnrotot & ","
                                StrSql = StrSql & rsEstructura!tenro & "," & tercero & ",1"
                            End If
                            monto_saldo = (monto_total - cubvalor1 - monto_saldo)
                            cant_saldo = (cant_total - cubvalor2 - cant_saldo)
                            monto_total = 0
                            cant_total = 0
                                    
                            'Para que no quede saldo cuando proporciona
                            If monto_saldo <= 1 And monto_saldo > 0 Then
                                cubvalor1 = cubvalor1 + monto_saldo
                            End If
                            If cant_saldo <= 1 And cant_saldo > 0 Then
                                'cubvalor2 = cubvalor2 + cant_saldo
                            End If
                            
                            'Si existe el cubo entonces actualizo
                            If Not rs.EOF Then
                                StrSql = "UPDATE anrcubo SET" & _
                                    " anrcubvalor1 = " & rs!anrcubvalor1 * cubvalor1 & _
                                    " ,anrcubvalor2 = " & rs!anrcubvalor2 * cubvalor2
                                    '---------------------------
                                    ' FAF 14-02-2005
                                    If CInt(rsFactor!facpresup) = -1 Then
                                        cubvalor3 = CalcularPresupuestado(rs!anrcubvalor1 * cubvalor1, rsFactor!facpresup, rsFactor!facopfijo, rsFactor!facopsuma, rsFactor!facpresupmonto)
                                        cubvalor4 = CalcularPresupuestado(rs!anrcubvalor2 * cubvalor2, rsFactor!facpresup, rsFactor!facopfijo, rsFactor!facopsuma, rsFactor!facpresupmonto)
                                        StrSql = StrSql & " ,anrcubvalor3 = " & cubvalor3
                                        StrSql = StrSql & " ,anrcubvalor4 = " & cubvalor4
                                    End If
                                    '---------------------------
                                StrSql = StrSql & " WHERE anrcabnro = " & rsAnrCab!anrcabnro & _
                                    " AND facnro = " & rsFactor!facnrotot & _
                                    " AND tenro = " & rsEstructura!tenro & _
                                    " AND estrnro = " & rsEstructura!estrnro & _
                                    " AND ternro = " & tercero & _
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
                        'Si no proporciona tomo al 100% y la ultima his_estruc del rango
                        porcentaje = 100
                        'If Last_OF_Factor() Or Last_OF_estrnro() Then
                        If Ultimo(rs_AnrCubo) Or Last_OF_estrnro() Then
                            If Not Last_OF_tenro() Then
                                monto_total = 0
                                cant_total = 0
                            Else
                                'Busco la ultima his_estr dentro del rango
                                StrSql = "SELECT * FROM his_estructura " & _
                                    " WHERE his_estructura.ternro = " & tercero & _
                                    " AND his_estructura.tenro = " & rsEstructura!tenro & _
                                    " AND his_estructura.htetdesde <= " & ConvFecha(Fin_Per_Analizado) & _
                                    " AND (his_estructura.htethasta >= " & ConvFecha(Inicio_Per_Analizado) & _
                                    " OR his_estructura.htethasta IS NULL) "
                                OpenRecordset StrSql, objRs
                                objRs.MoveLast
                                
                                If Not objRs.EOF Then
                                    StrSql = "SELECT * FROM anrcubo" & _
                                        " WHERE anrcabnro = " & rsAnrCab!anrcabnro & _
                                        " AND facnro = " & rsFactor!facnrotot & _
                                        " AND tenro = " & objRs!tenro & _
                                        " AND estrnro = " & objRs!estrnro & _
                                        " AND ternro = " & tercero & _
                                        " AND anrrangnro = " & rsRango!anrrangnro
                                    OpenRecordset StrSql, rs
                                    
                                    cubvalor1 = monto_total * porcentaje / 100
                                    cubvalor2 = cant_total * porcentaje / 100
                                    
                                    If rs.EOF Then
                                        'Creo el cubo
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
                                            objRs!estrnro & "," & rsFactor!facnrotot & "," & _
                                            objRs!tenro & "," & tercero & ",1" & _
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
                                            Aux_cubvalor1 = cubvalor1
                                            Aux_cubvalor2 = cubvalor2
                                    Else
                                        StrSql = "UPDATE anrcubo SET" & _
                                            " anrcubvalor1 = " & rs!anrcubvalor1 * cubvalor1 & _
                                            " ,anrcubvalor2 = " & rs!anrcubvalor2 * cubvalor2
                                            '---------------------------
                                            ' FAF 14-02-2005
                                            If CInt(rsFactor!facpresup) = -1 Then
                                                cubvalor3 = CalcularPresupuestado(rs!anrcubvalor1 * cubvalor1, rsFactor!facpresup, rsFactor!facopfijo, rsFactor!facopsuma, rsFactor!facpresupmonto)
                                                cubvalor4 = CalcularPresupuestado(rs!anrcubvalor2 * cubvalor2, rsFactor!facpresup, rsFactor!facopfijo, rsFactor!facopsuma, rsFactor!facpresupmonto)
                                                StrSql = StrSql & " ,anrcubvalor3 = " & cubvalor3
                                                StrSql = StrSql & " ,anrcubvalor4 = " & cubvalor4
                                            End If
                                            '---------------------------
                                        StrSql = StrSql & " WHERE anrcabnro = " & rsAnrCab!anrcabnro & _
                                            " AND facnro = " & rsFactor!facnrotot & _
                                            " AND tenro = " & objRs!tenro & _
                                            " AND estrnro = " & objRs!estrnro & _
                                            " AND ternro = " & tercero & _
                                            " AND anrrangnro = " & rsRango!anrrangnro
                                            Aux_cubvalor1 = rs!anrcubvalor1 * cubvalor1
                                            Aux_cubvalor2 = rs!anrcubvalor2 * cubvalor2
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
            
                rsRango.MoveNext
            Loop
                
            ' Actualizo el progreso
            Progreso = Progreso + IncPorc
            StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & " WHERE bpronro = " & NroProceso
            objconnProgreso.Execute StrSql, , adExecuteNoRecords
                
            rs_AnrCubo.MoveNext
        Loop
        
siguientelegajo:
        rsFiltro.MoveNext
    Loop
    rsFactor.MoveNext
Loop
SumPorcTiempo = CDbl(SumPorcTiempo) + CDbl(PorcTiempo)
Exit Sub
CE:
    HuboErrorTipo = True
    HuboError = True
    Flog.writeline Espacios(Tabulador * 1) & "Error " & Err.Description
End Sub



Public Sub RestaFactores(Nro_Analisis As Long, Filtrar As Boolean)
' ---------------------------------------------------------------------------------------------
' Descripcion: Analisis para tipo de factor 9 (Resta de Factores)
' Autor      : FGZ
' Fecha      : 01/12/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
'Variables locales
Dim cant_flt As Long
Dim Desde As Date
Dim Hasta As Date
Dim horas As Single
Dim NroCab As Long
Dim perpago_desde As Long
Dim perpago_hasta As Long

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
Dim Aux_cubvalor1 As Single
Dim Aux_cubvalor2 As Single


'Variables para los first y last
Dim PrimerFactOri As Boolean
Dim TipoEstr As Long
Dim EstrAct As Long
Dim FactOri As Long
Dim MiConcepto As String

Dim estr_liqNex As String
Dim cod_cptoNex As String

' Fechas desde y hasta a analizar por el proceso entero
Dim Inicio_Per_Analizado As Date
Dim Fin_Per_Analizado As Date

' Fechas parciales que se estan analizando
Dim Dia_Inicio_Per_Analizado As Date
Dim Dia_Fin_Per_Analizado As Date

' Auxiliares para el manejo de ls ciclos
Dim MesActual As Integer
Dim MesInicio As Integer
Dim MesFin As Integer
Dim AnioInicio As Integer
Dim AnioFin As Integer
Dim AnioActual As Integer
Dim AuxDia As Integer
Dim Ok As Boolean

Dim UltimoDiaMes As Integer
Dim AuxRangoDesde As Date
Dim AuxRangoHasta As Date

'FGZ - 11/09/2003
Dim NombreBD As String

Dim rs As New ADODB.Recordset
Dim rs_AnrCubo As New ADODB.Recordset

'Código -------------------------------------------------------------------
'Abro la conexion para nexus
'OpenConnection strConexionNexus, objConnNexus

On Error GoTo CE:

' --- fgz 07/07/2003---------
'Obtengo la cabecera y el filtro
Call ObtenerCabecerayFiltro(Nro_Analisis, rsAnrCab, Filtrar, rsFiltro, cant_flt, Ok)

If Not Ok Then
    Exit Sub
End If
CantFiltro = 0
If Not rsFiltro.EOF Then
    'CantFiltro = rsFiltro.RecordCount
    CantFiltro = cant_flt
Else
    CantFiltro = 1
End If


'Estoy reprocesando
Call PurgarCubo(Nro_Analisis, 9)
' ---------------------------

'Comienzo el procesamiento
'StrSql = "SELECT * FROM anrcab_fact" & _
'    " INNER JOIN anrfact_ori ON anrfact_ori.facnro = anrcab_fact.facnro" & _
'    " AND anrfact_ori.tipfacnro = 8" & _
'    " INNER JOIN anrfactor ON anrfactor.facnro = anrcab_fact.facnro" & _
'    " WHERE anrcabnro = " & rsAnrCab!anrcabnro & _
'    " ORDER BY anrfact_ori.facnro"
StrSql = "SELECT anrcab_fact.facnro, anrcab_fact.anrcabnro, anrfactor.tipfacnro, anrfactor.facpropor, "
StrSql = StrSql & " anrfact_tot.facnrotot, anrfact_tot.anrfactsuma, anrfactor.facpresup, anrfactor.facpresupmonto, anrfactor.facopsuma, anrfactor.facopfijo "
StrSql = StrSql & " FROM anrcab_fact"
StrSql = StrSql & " INNER JOIN anrfact_tot ON anrfact_tot.facnro = anrcab_fact.facnro"
StrSql = StrSql & " INNER JOIN anrfactor ON anrfact_tot.facnrotot = anrfactor.facnro"
StrSql = StrSql & " WHERE anrcab_fact.anrcabnro = " & Nro_Analisis
StrSql = StrSql & " AND anrfactor.tipfacnro = 9"
StrSql = StrSql & " ORDER BY anrcab_fact.nrocolum, anrcab_fact.facnro"
OpenRecordset StrSql, rsFactor
    
CantFactor = 0
If Not rsFactor.EOF Then
    CantFactor = rsFactor.RecordCount
Else
    CantFactor = 1
End If

'obtengo el conjunto de legajos a procesar
Inicio_Per_Analizado = rsAnrCab!anrcabfecdesde
Fin_Per_Analizado = rsAnrCab!anrcabfechasta
Call ObtenerLegajos(2, Filtrar, rsAnrCab!anrcabnro, rsFiltro, rsAnrCab!anrcabfecdesde, rsAnrCab!anrcabfechasta)

'Seteo el incremento del progreso
'Progreso = 95
Progreso = SumPorcTiempo
If (CantFactor * rsFiltro.RecordCount * CantFiltro) <> 0 Then
    IncPorc = PorcTiempo / (CantFactor * rsFiltro.RecordCount)
Else
    IncPorc = PorcTiempo
End If
                    
Do While Not rsFactor.EOF
'    'Busco si es factor totalizador
'    StrSql = "SELECT * FROM anrfact_tot" & _
'            " WHERE facnro = " & FactOri
'    OpenRecordset StrSql, rsFactorTotalizador
'    If Not rsFactorTotalizador.EOF Then
'        Totaliza = True
'        ' codigo de factor con el cual se inserta en el cubo
'        FactorTotalizador = rsFactorTotalizador!facnrotot
'    Else
'        Totaliza = False
'        FactorTotalizador = 0
'    End If
    
    'Primer legajo del filtro
    rsFiltro.MoveFirst
    
    Do While Not rsFiltro.EOF
        tercero = rsFiltro!Ternro
        If (cant_flt > 0) And (rsFiltro!cant_te < cant_flt) Then
            GoTo siguientelegajo
        End If
        'Cuando hay establecido un filtro, se debe verificar que el empleado verifique
        'todos los filtros en el intervalo de tiempo analizado. El control se hace de
        'esta forma, para considerar en forma correcta los casos en donde existe más de
        'un registro en HIS_ESTRUCTURA, para el mismo empleado y para el mismo tipo de
        'estructura, que satisfacen el intervalo de tiempo.
    
    
        ' Recorre para el analisis las tablas segun los factores configurados
        StrSql = "SELECT DISTINCT ternro, facnro, anrcubvalor1, anrcubvalor2, anrrangnro FROM anrcubo"
        StrSql = StrSql & " WHERE anrcabnro = " & Nro_Analisis
        StrSql = StrSql & " AND facnro = " & rsFactor!facnro
'        StrSql = StrSql & " AND tenro = " & rsEstructura!tenro
'        StrSql = StrSql & " AND estrnro = " & rsEstructura!estrnro
        StrSql = StrSql & " AND ternro = " & tercero
'        StrSql = StrSql & " AND anrrangnro = " & rsRango!anrrangnro
        OpenRecordset StrSql, rs_AnrCubo
        
        Do While Not rs_AnrCubo.EOF
                    
            StrSql = "SELECT * FROM anrrangofec"
            StrSql = StrSql & " WHERE anrrangofec.anrcabnro = " & rsAnrCab!anrcabnro
            StrSql = StrSql & " AND anrrangorepro = -1"
            StrSql = StrSql & " AND anrrangnro = " & rs_AnrCubo!anrrangnro
'            StrSql = StrSql & " AND anrrangofec.anrrangfecdesde <= " & ConvFecha(rs_detliq!profecini)
'            StrSql = StrSql & " AND anrrangofec.anrrangfechasta >= " & ConvFecha(rs_detliq!profecfin)
            OpenRecordset StrSql, rsRango
            
            Do While Not rsRango.EOF
                Fin_Per_Analizado = rsRango!anrrangfechasta
                Inicio_Per_Analizado = rsRango!anrrangfecdesde
                    
                ' FGZ 10/07/2003--------------------------
                Call ObtenerEstructuras(Filtrar, tercero, Inicio_Per_Analizado, Fin_Per_Analizado, rsEstructura)
                If Not rsEstructura.EOF Then
                    TipoEstr = rsEstructura!tenro
                    EstrAct = rsEstructura!estrnro
                End If
            
                Do While Not rsEstructura.EOF
                    If PrimerFactOri Then
                        cantdiasper = DateDiff("d", Inicio_Per_Analizado, Fin_Per_Analizado) + 1
                        monto_total = 0
                        cant_total = 0
                        cant_saldo = 0
                        PrimerFactOri = False
                    End If
                                           
                    'Acumulo por Factor
                    'If CBool(rsFactor!anrfactsuma) Then
                        monto_total = monto_total + rs_AnrCubo!anrcubvalor1
                        cant_total = cant_total + rs_AnrCubo!anrcubvalor2
                    'Else
                    '    monto_total = monto_total + rs_AnrCubo!anrcubvalor1
                    '    cant_total = cant_total + rs_AnrCubo!anrcubvalor2
                    'End If
                    
                    'Calculo los dias de rango entre las fechas del rango y el his_estruct para proporcionar
                    If rsFactor!facpropor = -1 Then
                        If rsEstructura!htetdesde < Inicio_Per_Analizado Then
                            If rsEstructura!htethasta < Fin_Per_Analizado And (Not IsNull(rsEstructura!htethasta)) Then
                                cantdiasran = DateDiff("d", Inicio_Per_Analizado, rsEstructura!htethasta) + 1
                            Else
                                cantdiasran = DateDiff("d", Inicio_Per_Analizado, Fin_Per_Analizado) + 1
                            End If
                        Else
                            If (rsEstructura!htethasta < Fin_Per_Analizado) And (Not IsNull(rsEstructura!htethasta)) Then
                                cantdiasran = DateDiff("d", rsEstructura!htetdesde, rsEstructura!htethasta) + 1
                            Else
                                cantdiasran = DateDiff("d", rsEstructura!htetdesde, Fin_Per_Analizado) + 1
                            End If
                        End If
                                
                        'Porcentaje segun la cant. de dias en la his_estrutura
                        porcentaje = cantdiasran * 100 / cantdiasper
                                
                        'If Last_OF_Factor() Or Last_OF_estrnro() Then
                        If Ultimo(rs_AnrCubo) Or Last_OF_estrnro() Then
                            cubvalor1 = monto_total * porcentaje / 100
                            cubvalor2 = cant_total * porcentaje / 100
                                    
                            StrSql = "SELECT * FROM anrcubo"
                            StrSql = StrSql & " WHERE anrcabnro = " & Nro_Analisis
                            StrSql = StrSql & " AND facnro = " & rsFactor!facnrotot
                            StrSql = StrSql & " AND tenro = " & rsEstructura!tenro
                            StrSql = StrSql & " AND estrnro = " & rsEstructura!estrnro
                            StrSql = StrSql & " AND ternro = " & tercero
                            StrSql = StrSql & " AND anrrangnro = " & rsRango!anrrangnro
                            OpenRecordset StrSql, rs
        
                            'Si el cubo no existe lo creo
                            If rs.EOF Then
                                'Creo el cubo
                                StrSql = "INSERT INTO anrcubo(anrcabnro,anrcubmanual"
                                StrSql = StrSql & ",anrrangnro,estrnro,facnro,tenro,Ternro,tipnro"
                                StrSql = StrSql & ",anrcubvalor1,anrcubvalor2"
                                '---------------------------
                                ' FAF 14-02-2005
                                If CInt(rsFactor!facpresup) = -1 Then
                                    StrSql = StrSql & ",anrcubvalor3,anrcubvalor4"
                                End If
                                '---------------------------
                                StrSql = StrSql & ") VALUES ("
                                StrSql = StrSql & Nro_Analisis & ",0," & rsRango!anrrangnro & ","
                                StrSql = StrSql & rsEstructura!estrnro & "," & rsFactor!facnrotot & ","
                                StrSql = StrSql & rsEstructura!tenro & "," & tercero & ",1"
                            End If
                            monto_saldo = (monto_total - cubvalor1 - monto_saldo)
                            cant_saldo = (cant_total - cubvalor2 - cant_saldo)
                            monto_total = 0
                            cant_total = 0
                                    
                            'Para que no quede saldo cuando proporciona
                            If monto_saldo <= 1 And monto_saldo > 0 Then
                                cubvalor1 = cubvalor1 + monto_saldo
                            End If
                            If cant_saldo <= 1 And cant_saldo > 0 Then
                                'cubvalor2 = cubvalor2 + cant_saldo
                            End If
                                       
                            'Si existe el cubo entonces actualizo
                            If Not rs.EOF Then
                                If CBool(rsFactor!anrfactsuma) Then
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
                                        " AND facnro = " & rsFactor!facnrotot & _
                                        " AND tenro = " & rsEstructura!tenro & _
                                        " AND estrnro = " & rsEstructura!estrnro & _
                                        " AND ternro = " & tercero & _
                                        " AND anrrangnro = " & rsRango!anrrangnro
                                Else
                                    StrSql = "UPDATE anrcubo SET" & _
                                        " anrcubvalor1 = " & rs!anrcubvalor1 - cubvalor1 & _
                                        " ,anrcubvalor2 = " & rs!anrcubvalor2 - cubvalor2
                                        '---------------------------
                                        ' FAF 14-02-2005
                                        If CInt(rsFactor!facpresup) = -1 Then
                                            cubvalor3 = CalcularPresupuestado(rs!anrcubvalor1 - cubvalor1, rsFactor!facpresup, rsFactor!facopfijo, rsFactor!facopsuma, rsFactor!facpresupmonto)
                                            cubvalor4 = CalcularPresupuestado(rs!anrcubvalor2 - cubvalor2, rsFactor!facpresup, rsFactor!facopfijo, rsFactor!facopsuma, rsFactor!facpresupmonto)
                                            StrSql = StrSql & " ,anrcubvalor3 = " & cubvalor3
                                            StrSql = StrSql & " ,anrcubvalor4 = " & cubvalor4
                                        End If
                                        '---------------------------
                                    StrSql = StrSql & " WHERE anrcabnro = " & rsAnrCab!anrcabnro & _
                                        " WHERE anrcabnro = " & rsAnrCab!anrcabnro & _
                                        " AND facnro = " & rsFactor!facnrotot & _
                                        " AND tenro = " & rsEstructura!tenro & _
                                        " AND estrnro = " & rsEstructura!estrnro & _
                                        " AND ternro = " & tercero & _
                                        " AND anrrangnro = " & rsRango!anrrangnro
                                End If
                            Else
                                StrSql = StrSql & "," & IIf(CBool(rsFactor!anrfactsuma), cubvalor1, cubvalor1 * -1) & "," & IIf(CBool(rsFactor!anrfactsuma), cubvalor2, cubvalor2 * -1)
                                    '---------------------------
                                    ' FAF 14-02-2005
                                    If CInt(rsFactor!facpresup) = -1 Then
                                        cubvalor3 = CalcularPresupuestado(cubvalor1, rsFactor!facpresup, rsFactor!facopfijo, rsFactor!facopsuma, rsFactor!facpresupmonto)
                                        cubvalor4 = CalcularPresupuestado(cubvalor2, rsFactor!facpresup, rsFactor!facopfijo, rsFactor!facopsuma, rsFactor!facpresupmonto)
                                        StrSql = StrSql & "," & IIf(CBool(rsFactor!anrfactsuma), cubvalor3, cubvalor3 * -1) & "," & IIf(CBool(rsFactor!anrfactsuma), cubvalor4, cubvalor4 * -1)
                                    End If
                                    '---------------------------
                                StrSql = StrSql & ")"
                            End If
                            objConn.Execute StrSql, , adExecuteNoRecords
                                    
                        End If
                    Else
                        'Si no proporciona tomo al 100% y la ultima his_estruc del rango
                        porcentaje = 100
                        'If Last_OF_Factor() Or Last_OF_estrnro() Then
                        If Ultimo(rs_AnrCubo) Or Last_OF_estrnro() Then
                            If Not Last_OF_tenro() Then
                                monto_total = 0
                                cant_total = 0
                            Else
                                'Busco la ultima his_estr dentro del rango
                                StrSql = "SELECT * FROM his_estructura " & _
                                    " WHERE his_estructura.ternro = " & tercero & _
                                    " AND his_estructura.tenro = " & rsEstructura!tenro & _
                                    " AND his_estructura.htetdesde <= " & ConvFecha(Fin_Per_Analizado) & _
                                    " AND (his_estructura.htethasta >= " & ConvFecha(Inicio_Per_Analizado) & _
                                    " OR his_estructura.htethasta IS NULL) "
                                OpenRecordset StrSql, objRs
                                objRs.MoveLast
                                
                                If Not objRs.EOF Then
                                    StrSql = "SELECT * FROM anrcubo" & _
                                        " WHERE anrcabnro = " & rsAnrCab!anrcabnro & _
                                        " AND facnro = " & rsFactor!facnrotot & _
                                        " AND tenro = " & objRs!tenro & _
                                        " AND estrnro = " & objRs!estrnro & _
                                        " AND ternro = " & tercero & _
                                        " AND anrrangnro = " & rsRango!anrrangnro
                                    OpenRecordset StrSql, rs
                                    
                                    cubvalor1 = monto_total * porcentaje / 100
                                    cubvalor2 = cant_total * porcentaje / 100
                                    
                                    If rs.EOF Then
                                        'Creo el cubo
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
                                            objRs!estrnro & "," & rsFactor!facnrotot & "," & _
                                            objRs!tenro & "," & tercero & ",1" & _
                                            "," & IIf(CBool(rsFactor!anrfactsuma), cubvalor1, cubvalor1 * -1) & "," & IIf(CBool(rsFactor!anrfactsuma), cubvalor2, cubvalor2 * -1)
                                            '---------------------------
                                            ' FAF 14-02-2005
                                            If CInt(rsFactor!facpresup) = -1 Then
                                                cubvalor3 = CalcularPresupuestado(cubvalor1, rsFactor!facpresup, rsFactor!facopfijo, rsFactor!facopsuma, rsFactor!facpresupmonto)
                                                cubvalor4 = CalcularPresupuestado(cubvalor2, rsFactor!facpresup, rsFactor!facopfijo, rsFactor!facopsuma, rsFactor!facpresupmonto)
                                                StrSql = StrSql & "," & IIf(CBool(rsFactor!anrfactsuma), cubvalor3, cubvalor3 * -1) & "," & IIf(CBool(rsFactor!anrfactsuma), cubvalor4, cubvalor4 * -1)
                                            End If
                                            '---------------------------
                                        StrSql = StrSql & ")"
                                            Aux_cubvalor1 = IIf(CBool(rsFactor!anrfactsuma), cubvalor1, cubvalor1 * -1)
                                            Aux_cubvalor2 = IIf(CBool(rsFactor!anrfactsuma), cubvalor2, cubvalor2 * -1)
                                    Else
                                        If CBool(rsFactor!anrfactsuma) Then
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
                                                " AND facnro = " & rsFactor!facnrotot & _
                                                " AND tenro = " & objRs!tenro & _
                                                " AND estrnro = " & objRs!estrnro & _
                                                " AND ternro = " & tercero & _
                                                " AND anrrangnro = " & rsRango!anrrangnro
                                                Aux_cubvalor1 = rs!anrcubvalor1 + cubvalor1
                                                Aux_cubvalor2 = rs!anrcubvalor2 + cubvalor2
                                        Else
                                            StrSql = "UPDATE anrcubo SET" & _
                                                " anrcubvalor1 = " & rs!anrcubvalor1 - cubvalor1 & _
                                                " ,anrcubvalor2 = " & rs!anrcubvalor2 - cubvalor2
                                                '---------------------------
                                                ' FAF 14-02-2005
                                                If CInt(rsFactor!facpresup) = -1 Then
                                                    cubvalor3 = CalcularPresupuestado(rs!anrcubvalor1 - cubvalor1, rsFactor!facpresup, rsFactor!facopfijo, rsFactor!facopsuma, rsFactor!facpresupmonto)
                                                    cubvalor4 = CalcularPresupuestado(rs!anrcubvalor2 - cubvalor2, rsFactor!facpresup, rsFactor!facopfijo, rsFactor!facopsuma, rsFactor!facpresupmonto)
                                                    StrSql = StrSql & " ,anrcubvalor3 = " & cubvalor3
                                                    StrSql = StrSql & " ,anrcubvalor4 = " & cubvalor4
                                                End If
                                                '---------------------------
                                            StrSql = StrSql & " WHERE anrcabnro = " & rsAnrCab!anrcabnro & _
                                                " AND facnro = " & rsFactor!facnrotot & _
                                                " AND tenro = " & objRs!tenro & _
                                                " AND estrnro = " & objRs!estrnro & _
                                                " AND ternro = " & tercero & _
                                                " AND anrrangnro = " & rsRango!anrrangnro
                                                Aux_cubvalor1 = rs!anrcubvalor1 - cubvalor1
                                                Aux_cubvalor2 = rs!anrcubvalor2 - cubvalor2
                                        End If
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
            
                rsRango.MoveNext
            Loop
                
            ' Actualizo el progreso
            Progreso = Progreso + IncPorc
            StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & " WHERE bpronro = " & NroProceso
            objconnProgreso.Execute StrSql, , adExecuteNoRecords
                
            rs_AnrCubo.MoveNext
        Loop
        
siguientelegajo:
        rsFiltro.MoveNext
    Loop
    rsFactor.MoveNext
Loop
SumPorcTiempo = CDbl(SumPorcTiempo) + CDbl(PorcTiempo)
Exit Sub
CE:
    HuboErrorTipo = True
    HuboError = True
    Flog.writeline Espacios(Tabulador * 1) & "Error " & Err.Description
End Sub


Public Sub DivideFactores(Nro_Analisis As Long, Filtrar As Boolean)
' ---------------------------------------------------------------------------------------------
' Descripcion: Analisis para tipo de factor 11 (Division de Factores)
' Autor      : FGZ
' Fecha      : 01/12/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
'Variables locales
Dim cant_flt As Long
Dim Desde As Date
Dim Hasta As Date
Dim horas As Single
Dim NroCab As Long
Dim perpago_desde As Long
Dim perpago_hasta As Long

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
Dim Aux_cubvalor1 As Single
Dim Aux_cubvalor2 As Single
Dim Aux_cubvalor3 As Single
Dim Aux_cubvalor4 As Single


'Variables para los first y last
Dim PrimerFactOri As Boolean
Dim TipoEstr As Long
Dim EstrAct As Long
Dim FactOri As Long
Dim MiConcepto As String

Dim estr_liqNex As String
Dim cod_cptoNex As String

' Fechas desde y hasta a analizar por el proceso entero
Dim Inicio_Per_Analizado As Date
Dim Fin_Per_Analizado As Date

' Fechas parciales que se estan analizando
Dim Dia_Inicio_Per_Analizado As Date
Dim Dia_Fin_Per_Analizado As Date

' Auxiliares para el manejo de ls ciclos
Dim MesActual As Integer
Dim MesInicio As Integer
Dim MesFin As Integer
Dim AnioInicio As Integer
Dim AnioFin As Integer
Dim AnioActual As Integer
Dim AuxDia As Integer
Dim Ok As Boolean

Dim UltimoDiaMes As Integer
Dim AuxRangoDesde As Date
Dim AuxRangoHasta As Date

'FGZ - 11/09/2003
Dim NombreBD As String

Dim rs As New ADODB.Recordset
Dim rs_AnrCubo As New ADODB.Recordset

'Código -------------------------------------------------------------------
'Abro la conexion para nexus
'OpenConnection strConexionNexus, objConnNexus

On Error GoTo CE:

' --- fgz 07/07/2003---------
'Obtengo la cabecera y el filtro
Call ObtenerCabecerayFiltro(Nro_Analisis, rsAnrCab, Filtrar, rsFiltro, cant_flt, Ok)

If Not Ok Then
    Exit Sub
End If
CantFiltro = 0
If Not rsFiltro.EOF Then
    'CantFiltro = rsFiltro.RecordCount
    CantFiltro = cant_flt
Else
    CantFiltro = 1
End If


'Estoy reprocesando
Call PurgarCubo(Nro_Analisis, 11)
' ---------------------------

'Comienzo el procesamiento
'StrSql = "SELECT * FROM anrcab_fact" & _
'    " INNER JOIN anrfact_ori ON anrfact_ori.facnro = anrcab_fact.facnro" & _
'    " AND anrfact_ori.tipfacnro = 8" & _
'    " INNER JOIN anrfactor ON anrfactor.facnro = anrcab_fact.facnro" & _
'    " WHERE anrcabnro = " & rsAnrCab!anrcabnro & _
'    " ORDER BY anrfact_ori.facnro"
StrSql = "SELECT anrcab_fact.facnro, anrcab_fact.anrcabnro, anrfactor.tipfacnro, anrfactor.facpropor, "
StrSql = StrSql & " anrfact_tot.facnrotot, anrfact_tot.anrfactsuma, anrfactor.facpresup, anrfactor.facpresupmonto, anrfactor.facopsuma, anrfactor.facopfijo "
StrSql = StrSql & " FROM anrcab_fact"
StrSql = StrSql & " INNER JOIN anrfact_tot ON anrfact_tot.facnro = anrcab_fact.facnro"
StrSql = StrSql & " INNER JOIN anrfactor ON anrfact_tot.facnrotot = anrfactor.facnro"
StrSql = StrSql & " WHERE anrcab_fact.anrcabnro = " & Nro_Analisis
StrSql = StrSql & " AND anrfactor.tipfacnro = 11"
StrSql = StrSql & " ORDER BY anrcab_fact.nrocolum, anrcab_fact.facnro"
OpenRecordset StrSql, rsFactor
    
CantFactor = 0
If Not rsFactor.EOF Then
    CantFactor = rsFactor.RecordCount
Else
    CantFactor = 1
End If

'obtengo el conjunto de legajos a procesar
Inicio_Per_Analizado = rsAnrCab!anrcabfecdesde
Fin_Per_Analizado = rsAnrCab!anrcabfechasta
Call ObtenerLegajos(2, Filtrar, rsAnrCab!anrcabnro, rsFiltro, rsAnrCab!anrcabfecdesde, rsAnrCab!anrcabfechasta)

'Seteo el incremento del progreso
'Progreso = 95
Progreso = SumPorcTiempo
If (CantFactor * rsFiltro.RecordCount * CantFiltro) <> 0 Then
    IncPorc = PorcTiempo / (CantFactor * rsFiltro.RecordCount)
Else
    IncPorc = PorcTiempo
End If
                    
Do While Not rsFactor.EOF
'    'Busco si es factor totalizador
'    StrSql = "SELECT * FROM anrfact_tot" & _
'            " WHERE facnro = " & FactOri
'    OpenRecordset StrSql, rsFactorTotalizador
'    If Not rsFactorTotalizador.EOF Then
'        Totaliza = True
'        ' codigo de factor con el cual se inserta en el cubo
'        FactorTotalizador = rsFactorTotalizador!facnrotot
'    Else
'        Totaliza = False
'        FactorTotalizador = 0
'    End If
    
    'Primer legajo del filtro
    rsFiltro.MoveFirst
    
    Do While Not rsFiltro.EOF
        tercero = rsFiltro!Ternro
        If (cant_flt > 0) And (rsFiltro!cant_te < cant_flt) Then
            GoTo siguientelegajo
        End If
        'Cuando hay establecido un filtro, se debe verificar que el empleado verifique
        'todos los filtros en el intervalo de tiempo analizado. El control se hace de
        'esta forma, para considerar en forma correcta los casos en donde existe más de
        'un registro en HIS_ESTRUCTURA, para el mismo empleado y para el mismo tipo de
        'estructura, que satisfacen el intervalo de tiempo.
    
    
        ' Recorre para el analisis las tablas segun los factores configurados
        StrSql = "SELECT DISTINCT ternro, facnro, anrcubvalor1, anrcubvalor2, anrrangnro FROM anrcubo"
        StrSql = StrSql & " WHERE anrcabnro = " & Nro_Analisis
        StrSql = StrSql & " AND facnro = " & rsFactor!facnro
'        StrSql = StrSql & " AND tenro = " & rsEstructura!tenro
'        StrSql = StrSql & " AND estrnro = " & rsEstructura!estrnro
        StrSql = StrSql & " AND ternro = " & tercero
'        StrSql = StrSql & " AND anrrangnro = " & rsRango!anrrangnro
        OpenRecordset StrSql, rs_AnrCubo
        
        Do While Not rs_AnrCubo.EOF
                    
            StrSql = "SELECT * FROM anrrangofec"
            StrSql = StrSql & " WHERE anrrangofec.anrcabnro = " & rsAnrCab!anrcabnro
            StrSql = StrSql & " AND anrrangorepro = -1"
            StrSql = StrSql & " AND anrrangnro = " & rs_AnrCubo!anrrangnro
'            StrSql = StrSql & " AND anrrangofec.anrrangfecdesde <= " & ConvFecha(rs_detliq!profecini)
'            StrSql = StrSql & " AND anrrangofec.anrrangfechasta >= " & ConvFecha(rs_detliq!profecfin)
            OpenRecordset StrSql, rsRango
            
            Do While Not rsRango.EOF
                Fin_Per_Analizado = rsRango!anrrangfechasta
                Inicio_Per_Analizado = rsRango!anrrangfecdesde
                    
                ' FGZ 10/07/2003--------------------------
                Call ObtenerEstructuras(Filtrar, tercero, Inicio_Per_Analizado, Fin_Per_Analizado, rsEstructura)
                If Not rsEstructura.EOF Then
                    TipoEstr = rsEstructura!tenro
                    EstrAct = rsEstructura!estrnro
                End If
            
                Do While Not rsEstructura.EOF
                    If PrimerFactOri Then
                        cantdiasper = DateDiff("d", Inicio_Per_Analizado, Fin_Per_Analizado) + 1
                        monto_total = 0
                        cant_total = 0
                        cant_saldo = 0
                        PrimerFactOri = False
                    End If
                                           
                    'Acumulo por Factor
                    If CBool(rsFactor!anrfactsuma) Then
                        monto_total = monto_total + rs_AnrCubo!anrcubvalor1
                        cant_total = cant_total + rs_AnrCubo!anrcubvalor2
                    Else
                        If rs_AnrCubo!anrcubvalor1 <> 0 Then
                            monto_total = monto_total + rs_AnrCubo!anrcubvalor1
                        Else
                            monto_total = monto_total
                            Flog.writeline Espacios(Tabulador * 1) & "Division por 0. Monto del Factor " & rsFactor!facnro & " con valor 0"
                        End If
                        If rs_AnrCubo!anrcubvalor2 <> 0 Then
                            cant_total = cant_total + rs_AnrCubo!anrcubvalor2
                        Else
                            cant_total = cant_total
                            Flog.writeline Espacios(Tabulador * 1) & "Division por 0. Cantidad del Factor " & rsFactor!facnro & " con valor 0"
                        End If
                    End If
                    
                    'Calculo los dias de rango entre las fechas del rango y el his_estruct para proporcionar
                    If rsFactor!facpropor = -1 Then
                        If rsEstructura!htetdesde < Inicio_Per_Analizado Then
                            If rsEstructura!htethasta < Fin_Per_Analizado And (Not IsNull(rsEstructura!htethasta)) Then
                                cantdiasran = DateDiff("d", Inicio_Per_Analizado, rsEstructura!htethasta) + 1
                            Else
                                cantdiasran = DateDiff("d", Inicio_Per_Analizado, Fin_Per_Analizado) + 1
                            End If
                        Else
                            If (rsEstructura!htethasta < Fin_Per_Analizado) And (Not IsNull(rsEstructura!htethasta)) Then
                                cantdiasran = DateDiff("d", rsEstructura!htetdesde, rsEstructura!htethasta) + 1
                            Else
                                cantdiasran = DateDiff("d", rsEstructura!htetdesde, Fin_Per_Analizado) + 1
                            End If
                        End If
                                
                        'Porcentaje segun la cant. de dias en la his_estrutura
                        porcentaje = cantdiasran * 100 / cantdiasper
                                
                        If Last_OF_Factor() Or Last_OF_estrnro() Then
                            cubvalor1 = monto_total * porcentaje / 100
                            cubvalor2 = cant_total * porcentaje / 100
                                    
                            StrSql = "SELECT * FROM anrcubo"
                            StrSql = StrSql & " WHERE anrcabnro = " & Nro_Analisis
                            StrSql = StrSql & " AND facnro = " & rsFactor!facnrotot
                            StrSql = StrSql & " AND tenro = " & rsEstructura!tenro
                            StrSql = StrSql & " AND estrnro = " & rsEstructura!estrnro
                            StrSql = StrSql & " AND ternro = " & tercero
                            StrSql = StrSql & " AND anrrangnro = " & rsRango!anrrangnro
                            OpenRecordset StrSql, rs
        
                            'Si el cubo no existe lo creo
                            If rs.EOF Then
                                'Creo el cubo
                                StrSql = "INSERT INTO anrcubo(anrcabnro,anrcubmanual"
                                StrSql = StrSql & ",anrrangnro,estrnro,facnro,tenro,Ternro,tipnro"
                                StrSql = StrSql & ",anrcubvalor1,anrcubvalor2"
                                '---------------------------
                                ' FAF 14-02-2005
                                If CInt(rsFactor!facpresup) = -1 Then
                                    StrSql = StrSql & ",anrcubvalor3,anrcubvalor4"
                                End If
                                '---------------------------
                            StrSql = StrSql & ") VALUES (" & _
                                StrSql = StrSql & Nro_Analisis & ",0," & rsRango!anrrangnro & ","
                                StrSql = StrSql & rsEstructura!estrnro & "," & rsFactor!facnrotot & ","
                                StrSql = StrSql & rsEstructura!tenro & "," & tercero & ",1"
                            End If
                            monto_saldo = (monto_total - cubvalor1 - monto_saldo)
                            cant_saldo = (cant_total - cubvalor2 - cant_saldo)
                            monto_total = 0
                            cant_total = 0
                                    
                            'Para que no quede saldo cuando proporciona
                            If monto_saldo <= 1 And monto_saldo > 0 Then
                                cubvalor1 = cubvalor1 + monto_saldo
                            End If
                            If cant_saldo <= 1 And cant_saldo > 0 Then
                                'cubvalor2 = cubvalor2 + cant_saldo
                            End If
                                       
                            'Si existe el cubo entonces actualizo
                            If Not rs.EOF Then
                                If CBool(rsFactor!anrfactsuma) Then
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
                                        " AND facnro = " & rsFactor!facnrotot & _
                                        " AND tenro = " & rsEstructura!tenro & _
                                        " AND estrnro = " & rsEstructura!estrnro & _
                                        " AND ternro = " & tercero & _
                                        " AND anrrangnro = " & rsRango!anrrangnro
                                Else
                                    StrSql = "UPDATE anrcubo SET"
                                    If cubvalor1 <> 0 Then
                                        StrSql = StrSql & " anrcubvalor1 = " & rs!anrcubvalor1 / cubvalor1
                                    Else
                                        Flog.writeline Espacios(Tabulador * 1) & "Division por 0. Monto del Factor " & rsFactor!facnro & " con valor 0"
                                        StrSql = StrSql & " anrcubvalor1 = " & rs!anrcubvalor1 / 1
                                    End If
                                    If cubvalor2 <> 0 Then
                                        StrSql = StrSql & " ,anrcubvalor2 = " & rs!anrcubvalor2 / cubvalor2
                                    Else
                                        Flog.writeline Espacios(Tabulador * 1) & "Division por 0. Cantidad del Factor " & rsFactor!facnro & " con valor 0"
                                        StrSql = StrSql & " ,anrcubvalor2 = " & rs!anrcubvalor2 / 1
                                    End If
                                    '---------------------------
                                    ' FAF 14-02-2005
                                    If CInt(rsFactor!facpresup) = -1 Then
                                        If cubvalor1 <> 0 Then
                                            cubvalor3 = CalcularPresupuestado(rs!anrcubvalor1 / cubvalor1, rsFactor!facpresup, rsFactor!facopfijo, rsFactor!facopsuma, rsFactor!facpresupmonto)
                                            StrSql = StrSql & " ,anrcubvalor3 = " & cubvalor3
                                        Else
                                            Flog.writeline Espacios(Tabulador * 1) & "Division por 0. Monto Presupuestado del Factor " & rsFactor!facnro & " con valor 0"
                                            cubvalor3 = CalcularPresupuestado(rs!anrcubvalor1 / 1, rsFactor!facpresup, rsFactor!facopfijo, rsFactor!facopsuma, rsFactor!facpresupmonto)
                                            StrSql = StrSql & " ,anrcubvalor3 = " & cubvalor3
                                        End If
                                        If cubvalor2 <> 0 Then
                                            cubvalor4 = CalcularPresupuestado(rs!anrcubvalor2 / cubvalor2, rsFactor!facpresup, rsFactor!facopfijo, rsFactor!facopsuma, rsFactor!facpresupmonto)
                                            StrSql = StrSql & " ,anrcubvalor4 = " & cubvalor4
                                        Else
                                            Flog.writeline Espacios(Tabulador * 1) & "Division por 0. Cantidad Presupuestada del Factor " & rsFactor!facnro & " con valor 0"
                                            cubvalor4 = CalcularPresupuestado(rs!anrcubvalor2 / 1, rsFactor!facpresup, rsFactor!facopfijo, rsFactor!facopsuma, rsFactor!facpresupmonto)
                                            StrSql = StrSql & " ,anrcubvalor4 = " & cubvalor4
                                        End If
                                    End If
                                    '---------------------------
                                    StrSql = StrSql & " WHERE anrcabnro = " & rsAnrCab!anrcabnro
                                    StrSql = StrSql & " AND facnro = " & rsFactor!facnrotot
                                    StrSql = StrSql & " AND tenro = " & rsEstructura!tenro
                                    StrSql = StrSql & " AND estrnro = " & rsEstructura!estrnro
                                    StrSql = StrSql & " AND ternro = " & tercero
                                    StrSql = StrSql & " AND anrrangnro = " & rsRango!anrrangnro
                                End If
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
                        'Si no proporciona tomo al 100% y la ultima his_estruc del rango
                        porcentaje = 100
                        'If Last_OF_Factor() Or Last_OF_estrnro() Then
                        If Ultimo(rs_AnrCubo) Or Last_OF_estrnro() Then
                            If Not Last_OF_tenro() Then
                                monto_total = 0
                                cant_total = 0
                            Else
                                'Busco la ultima his_estr dentro del rango
                                StrSql = "SELECT * FROM his_estructura " & _
                                    " WHERE his_estructura.ternro = " & tercero & _
                                    " AND his_estructura.tenro = " & rsEstructura!tenro & _
                                    " AND his_estructura.htetdesde <= " & ConvFecha(Fin_Per_Analizado) & _
                                    " AND (his_estructura.htethasta >= " & ConvFecha(Inicio_Per_Analizado) & _
                                    " OR his_estructura.htethasta IS NULL) "
                                OpenRecordset StrSql, objRs
                                objRs.MoveLast
                                
                                If Not objRs.EOF Then
                                    StrSql = "SELECT * FROM anrcubo" & _
                                        " WHERE anrcabnro = " & rsAnrCab!anrcabnro & _
                                        " AND facnro = " & rsFactor!facnrotot & _
                                        " AND tenro = " & objRs!tenro & _
                                        " AND estrnro = " & objRs!estrnro & _
                                        " AND ternro = " & tercero & _
                                        " AND anrrangnro = " & rsRango!anrrangnro
                                    OpenRecordset StrSql, rs
                                    
                                    cubvalor1 = monto_total * porcentaje / 100
                                    cubvalor2 = cant_total * porcentaje / 100
                                    
                                    If rs.EOF Then
                                        'Creo el cubo
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
                                            objRs!estrnro & "," & rsFactor!facnrotot & "," & _
                                            objRs!tenro & "," & tercero & ",1" & _
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
                                            Aux_cubvalor1 = cubvalor1
                                            Aux_cubvalor2 = cubvalor2
                                    Else
                                        If CBool(rsFactor!anrfactsuma) Then
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
                                                " AND facnro = " & rsFactor!facnrotot & _
                                                " AND tenro = " & objRs!tenro & _
                                                " AND estrnro = " & objRs!estrnro & _
                                                " AND ternro = " & tercero & _
                                                " AND anrrangnro = " & rsRango!anrrangnro
                                                Aux_cubvalor1 = rs!anrcubvalor1 + cubvalor1
                                                Aux_cubvalor2 = rs!anrcubvalor2 + cubvalor2
                                        Else
                                            StrSql = "UPDATE anrcubo SET"
                                            If cubvalor1 <> 0 Then
                                                StrSql = StrSql & " anrcubvalor1 = " & rs!anrcubvalor1 / cubvalor1
                                                Aux_cubvalor1 = rs!anrcubvalor1 / cubvalor1
                                            Else
                                                Flog.writeline Espacios(Tabulador * 1) & "Division por 0. Monto del Factor " & rsFactor!facnro & " con valor 0"
                                                StrSql = StrSql & " anrcubvalor1 = " & rs!anrcubvalor1 / 1
                                                Aux_cubvalor1 = rs!anrcubvalor1
                                            End If
                                            If cubvalor2 <> 0 Then
                                                StrSql = StrSql & " ,anrcubvalor2 = " & rs!anrcubvalor2 / cubvalor2
                                                Aux_cubvalor2 = rs!anrcubvalor2 / cubvalor2
                                            Else
                                                Flog.writeline Espacios(Tabulador * 1) & "Division por 0. Cantidad del Factor " & rsFactor!facnro & " con valor 0"
                                                StrSql = StrSql & " ,anrcubvalor2 = " & rs!anrcubvalor2 / 1
                                                Aux_cubvalor2 = rs!anrcubvalor2
                                            End If
                                            '---------------------------
                                            ' FAF 14-02-2005
                                            If CInt(rsFactor!facpresup) = -1 Then
                                                If cubvalor1 <> 0 Then
                                                    cubvalor3 = CalcularPresupuestado(rs!anrcubvalor1 / cubvalor1, rsFactor!facpresup, rsFactor!facopfijo, rsFactor!facopsuma, rsFactor!facpresupmonto)
                                                    StrSql = StrSql & " ,anrcubvalor3 = " & cubvalor3
                                                    Aux_cubvalor3 = cubvalor3
                                                Else
                                                    Flog.writeline Espacios(Tabulador * 1) & "Division por 0. Monto Presupuestado del Factor " & rsFactor!facnro & " con valor 0"
                                                    cubvalor3 = CalcularPresupuestado(rs!anrcubvalor1 / 1, rsFactor!facpresup, rsFactor!facopfijo, rsFactor!facopsuma, rsFactor!facpresupmonto)
                                                    StrSql = StrSql & " ,anrcubvalor3 = " & cubvalor3
                                                    Aux_cubvalor3 = cubvalor3
                                                End If
                                                If cubvalor2 <> 0 Then
                                                    cubvalor4 = CalcularPresupuestado(rs!anrcubvalor2 / cubvalor2, rsFactor!facpresup, rsFactor!facopfijo, rsFactor!facopsuma, rsFactor!facpresupmonto)
                                                    StrSql = StrSql & " ,anrcubvalor4 = " & cubvalor4
                                                    Aux_cubvalor4 = cubvalor4
                                                Else
                                                    Flog.writeline Espacios(Tabulador * 1) & "Division por 0. Cantidad Presupuestada del Factor " & rsFactor!facnro & " con valor 0"
                                                    cubvalor4 = CalcularPresupuestado(rs!anrcubvalor2 / 1, rsFactor!facpresup, rsFactor!facopfijo, rsFactor!facopsuma, rsFactor!facpresupmonto)
                                                    StrSql = StrSql & " ,anrcubvalor4 = " & cubvalor4
                                                    Aux_cubvalor4 = cubvalor4
                                                End If
                                            End If
                                            '---------------------------
                                            StrSql = StrSql & " WHERE anrcabnro = " & rsAnrCab!anrcabnro
                                            StrSql = StrSql & " AND facnro = " & rsFactor!facnrotot
                                            StrSql = StrSql & " AND tenro = " & objRs!tenro
                                            StrSql = StrSql & " AND estrnro = " & objRs!estrnro
                                            StrSql = StrSql & " AND ternro = " & tercero
                                            StrSql = StrSql & " AND anrrangnro = " & rsRango!anrrangnro
                                        End If
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
            
                rsRango.MoveNext
            Loop
                
            ' Actualizo el progreso
            Progreso = Progreso + IncPorc
            StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & " WHERE bpronro = " & NroProceso
            objconnProgreso.Execute StrSql, , adExecuteNoRecords
                
            rs_AnrCubo.MoveNext
        Loop
        
siguientelegajo:
        rsFiltro.MoveNext
    Loop
    rsFactor.MoveNext
Loop
SumPorcTiempo = CDbl(SumPorcTiempo) + CDbl(PorcTiempo)
Exit Sub
CE:
    HuboErrorTipo = True
    HuboError = True
    Flog.writeline Espacios(Tabulador * 1) & "Error " & Err.Description
End Sub


