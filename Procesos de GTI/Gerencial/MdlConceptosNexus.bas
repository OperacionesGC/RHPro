Attribute VB_Name = "MdlConceptosNexus"
Option Explicit

'Para Informix
Global Const strConexionNexus = "DSN=Nexushr-RHPro"
Global Const strConexionNexusConf = "DSN=NexushrConf-RHPro"
'11/09/2003 FGZ
Global Const NombreBDNexusConf = "HR3Conf"
Global Const NombreBDNexus = "HR3"

Global objConnNexus As New ADODB.Connection
Global objConnNexusConf As New ADODB.Connection


Public Sub ConceptosNexus(Nro_Analisis As Long, Filtrar As Boolean)
' ---------------------------------------------------------------------------------------------
' Descripcion: Analisis para tipo de factor 7 (Conceptos de Nexus)
' Autor      : FGZ
' Fecha      : 15/11/2004
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

'Variables para los first y last
Dim PrimerFactOri As Boolean
Dim TipoEstr As Long
Dim EstrAct As Long
Dim FactOri As Long
Dim MiConcepto As String

Dim rs As New ADODB.Recordset

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


'Código -------------------------------------------------------------------

'Abro la conexion para Nexus
OpenConnection strConexionNexus, objConnNexus
OpenConnection strConexionNexusConf, objConnNexusConf

' --- fgz 07/07/2003---------
'Obtengo la cabecera y el filtro
Call ObtenerCabecerayFiltro(Nro_Analisis, rsAnrCab, Filtrar, rsFiltro, cant_flt, Ok)

If Not Ok Then
    Exit Sub
End If

'Estoy reprocesando
Call PurgarCubo(Nro_Analisis, 7)
' ---------------------------

StrSql = "SELECT * FROM anrrangofec" & _
         " WHERE anrrangofec.anrcabnro = " & rsAnrCab!anrcabnro & _
         " AND anrrangorepro = -1 "
OpenRecordset StrSql, rsRango
                    
Do While Not rsRango.EOF
    ' separo el procesamiento en cada uno de los rangos definidos
    ' Recupero el inicio y fin del periodo a analizar
    
    Fin_Per_Analizado = rsRango!anrrangfechasta
    Inicio_Per_Analizado = rsRango!anrrangfecdesde

' divido en periodos por mes
    MesInicio = Month(Inicio_Per_Analizado)
    MesFin = Month(Fin_Per_Analizado)
    AnioInicio = Year(Inicio_Per_Analizado)
    AnioFin = Year(Fin_Per_Analizado)

    MesActual = MesInicio
    AnioActual = AnioInicio
    
Do While AnioActual <= AnioFin

    Do While (MesActual <= 12 And AnioActual < AnioFin) Or (MesActual <= MesFin And AnioActual <= AnioFin)
        ' dia de inicio del periodo a analizar
        If MesActual < 10 Then
            Dia_Inicio_Per_Analizado = CDate("01/0" & MesActual & "/" & AnioActual)
        Else
            Dia_Inicio_Per_Analizado = CDate("01/" & MesActual & "/" & AnioActual)
        End If
        
        ' dia de fin del periodo a analizar
        ' Ultimo dia del mes actual
        If MesActual <> 12 Then
            AuxDia = Day(CDate("01/" & MesActual + 1 & "/" & AnioActual) - 1)
        Else
            AuxDia = 31
        End If
        
        If MesActual < 10 Then
            Dia_Fin_Per_Analizado = CDate(AuxDia & "/0" & MesActual & "/" & AnioActual)
        Else
            Dia_Fin_Per_Analizado = CDate(AuxDia & "/" & MesActual & "/" & AnioActual)
        End If

        ' obtengo el conjunto de legajos a procesar
        Call ObtenerLegajos(1, Filtrar, rsAnrCab!anrcabnro, rsFiltro, Dia_Inicio_Per_Analizado, Dia_Fin_Per_Analizado)
    
        Progreso = SumPorcTiempo
        If Not rsFiltro.EOF Then
            IncPorc = 100 / rsFiltro.RecordCount
        End If
    
        ' para probar
       StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & " WHERE bpronro = " & NroProceso
       objConn.Execute StrSql, , adExecuteNoRecords
    
    
    
        perpago_desde = Year(Dia_Inicio_Per_Analizado)
        If (Month(Dia_Inicio_Per_Analizado) < 10) Then
          perpago_desde = perpago_desde & "0"
        End If
        perpago_desde = perpago_desde & Month(Dia_Inicio_Per_Analizado)
        
        perpago_hasta = Year(Dia_Fin_Per_Analizado)
        If (Month(Dia_Fin_Per_Analizado) < 10) Then
          perpago_hasta = perpago_hasta & "0"
        End If
        perpago_hasta = perpago_hasta & Month(Dia_Fin_Per_Analizado)

' Segun el intervalo de análisis, se determinan las cadenas en formato AAAAMM que son utilizadas
' para acceder a la tabla HISTLIQ. O.D.A. 30/06/2003
    
    Do While Not rsFiltro.EOF
      If (cant_flt > 0) And (rsFiltro!cant_te < cant_flt) Then
          GoTo siguientelegajo
      End If
' Cuando hay establecido un filtro, se debe verificar que el empleado verifique
' todos los filtros en el intervalo de tiempo analizado. El control se hace de
' esta forma, para considerar en forma correcta los casos en donde existe más de
' un registro en HIS_ESTRUCTURA, para el mismo empleado y para el mismo tipo de
' estructura, que satisfacen el intervalo de tiempo.
' Legajo 387589, tenro 36, Maciel, para Mayo 2003 O.D.A. 04/07/2003


' Recorre para el analisis las tablas de nexus segun los factores configurados
' Comienzo con las tablas de Nexus
        StrSql = " SELECT legaliq.periodo_pago, legaliq.nro_corr_liq, legaliq.nro_leg FROM histliq " & _
                 " INNER JOIN legaliq " & _
                 " ON    legaliq.periodo_pago   = histliq.periodo_pago" & _
                 " AND   legaliq.nro_corr_liq   = histliq.nro_corr_liq" & _
                 " WHERE histliq.liq_confirmada = 'S' " & _
                 " AND   histliq.periodo_pago   = '" & perpago_desde & "'" & _
                 " AND   legaliq.nro_leg        = " & rsFiltro!empleg & _
                 " GROUP BY legaliq.periodo_pago, legaliq.nro_corr_liq, legaliq.nro_leg " & _
                 " ORDER BY nro_leg"
                 
        ' FGZ - 15/08/2003
        ' si es legajo confidencial ==> utilizo otra conexion
        If UCase(rsFiltro!estrcodext) = "ROCA" Then
            OpenRecordsetNexusConf StrSql, rsHistliq
        Else
            OpenRecordsetNexus StrSql, rsHistliq
        End If
        
' Con esta forma de trabajo, se está accediendo a los procesos de liquidacion de un periodo
' en particular, por lo que no hay necesidad de tener un intervalo de periodos. O.D.A. 04/07/2003
        
        'FGZ - 11/09/2003
        If UCase(rsFiltro!estrcodext) = "ROCA" Then
            NombreBD = NombreBDNexusConf
        Else
            NombreBD = NombreBDNexus
        End If
        
        Do While Not rsHistliq.EOF
            'ODA - FGZ  11/09/2003
            ' desde la conexion de rhpro hago inner join con una tabla de Nexus
            ' Este fué el secreto de la mejora. (cambio de dos while anidados por unos solo)
            
            StrSql = "SELECT conc.conccod, conc.concabr, afact.*, afact_ori.*, acab_fact.*, hcon.* " & _
                " FROM anrcab_fact acab_fact " & _
                " INNER JOIN anrfact_ori afact_ori ON afact_ori.facnro = acab_fact.facnro " & _
                " AND afact_ori.tipfacnro = 7" & _
                " INNER JOIN anrfactor afact ON afact.facnro = acab_fact.facnro" & _
                " INNER JOIN concepto conc ON conc.concnro = afact_ori.faccodorig" & _
                " INNER JOIN " & NombreBD & ":histcon hcon ON hcon.periodo_pago = '" & rsHistliq!periodo_pago & "'" & _
                " AND   hcon.nro_corr_liq = " & rsHistliq!nro_corr_liq & _
                " AND   hcon.nro_leg      = " & rsHistliq!nro_leg & _
                " WHERE acab_fact.anrcabnro = " & rsAnrCab!anrcabnro & _
                " AND   hcon.estr_liq         = conc.conccod[1,8]" & _
                " AND   hcon.cod_cpto         = conc.conccod[9,12]" & _
                " ORDER BY hcon.nro_leg, afact_ori.facnro"
                
                OpenRecordset StrSql, rsHistCon
        
            If Not rsHistCon.EOF Then
                'Para el simular el first_of
                PrimerFactOri = True
                'Para el simular el last_of en la tabla anrfact_ori
                FactOri = rsHistCon!facnro
            End If
            
            Do While Not rsHistCon.EOF
            
'                '/* Busco si es factor totalizador
'                StrSql = "SELECT * FROM anrfact_tot" & _
'                        " WHERE facnro = " & rsHistCon!facnro
'                OpenRecordset StrSql, rsFactorTotalizador
'
'                If Not rsFactorTotalizador.EOF Then
'                    Totaliza = True
'                    ' codigo de factor con el cual se inserta en el cubo
'                    FactorTotalizador = rsFactorTotalizador!facnrotot
'                Else
'                    Totaliza = False
'                    FactorTotalizador = 0
'                End If
            
                    
                tercero = rsFiltro!Ternro
                    
             
                        ' FGZ 10/07/2003--------------------------
                        Call ObtenerEstructuras(Filtrar, tercero, Dia_Inicio_Per_Analizado, Dia_Fin_Per_Analizado, rsEstructura)
                        
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
                            monto_total = monto_total + rsHistCon!importe_final
                            cant_total = cant_total + (0 & rsHistCon!Cantidad)
                            
                            '/* Calculo los dias de rango entre las fechas del rango y
                            ' el his_estruct para proporcionar*/
                            If rsHistCon!facpropor = -1 Then
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
                                
                                If Last_OF_Factor() Or Last_OF_estrnro() Then
                                    
                                    cubvalor1 = monto_total * porcentaje / 100
                                    cubvalor2 = cant_total * porcentaje / 100
                                    
                                    StrSql = "SELECT * FROM anrcubo" & _
                                        " WHERE anrcabnro = " & rsAnrCab!anrcabnro & _
                                        " AND facnro = " & rsHistCon!facnro & _
                                        " AND tenro = " & rsEstructura!tenro & _
                                        " AND estrnro = " & rsEstructura!estrnro & _
                                        " AND ternro = " & tercero & _
                                        " AND anrrangnro = " & rsRango!anrrangnro
                                    OpenRecordset StrSql, rs
        
                                    'Si el cubo no existe lo creo
                                    If rs.EOF Then
                                    '/* Creo el cubo */
                                        StrSql = "INSERT INTO anrcubo(anrcabnro,anrcubmanual" & _
                                            ",anrrangnro,estrnro,facnro,tenro,Ternro,tipnro" & _
                                            ",anrcubvalor1,anrcubvalor2"
                                        StrSql = StrSql & ") VALUES (" & _
                                            rsAnrCab!anrcabnro & ",0," & rsRango!anrrangnro & "," & _
                                            rsEstructura!estrnro & "," & rsHistCon!facnro & "," & _
                                            rsEstructura!tenro & "," & tercero & ",1"
                                    End If
                                    
                                    monto_saldo = (monto_total - cubvalor1 - monto_saldo)
                                    cant_saldo = (cant_total - cubvalor2 - cant_saldo)
                                    monto_total = 0
                                    cant_total = 0
                                    
                                    '* Para que no quede saldo cuando proporciona */
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
                                        StrSql = StrSql & " WHERE anrcabnro = " & rsAnrCab!anrcabnro & _
                                            " AND facnro = " & rsHistCon!facnro & _
                                            " AND tenro = " & rsEstructura!tenro & _
                                            " AND estrnro = " & rsEstructura!estrnro & _
                                            " AND ternro = " & tercero & _
                                            " AND anrrangnro = " & rsRango!anrrangnro
                                    Else
                                        StrSql = StrSql & "," & cubvalor1 & "," & cubvalor2
                                        StrSql = StrSql & ")"
                                    End If
                                    
                                    objConn.Execute StrSql, , adExecuteNoRecords
                                    
'                                    'FZG 25/07/2003
'                                    'Actualizo Totalizador
'                                    If Totaliza Then
'                                        StrSql = "SELECT * FROM anrcubo" & _
'                                            " WHERE anrcabnro = " & rsAnrCab!anrcabnro & _
'                                            " AND facnro = " & FactorTotalizador & _
'                                            " AND tenro = " & rsEstructura!tenro & _
'                                            " AND estrnro = " & rsEstructura!estrnro & _
'                                            " AND ternro = " & tercero & _
'                                            " AND anrrangnro = " & rsRango!anrrangnro
'                                        OpenRecordset StrSql, rsTot
'
'                                        If rsTot.EOF Then
'                                            ' Creo el cubo
'                                            StrSql = "INSERT INTO anrcubo(anrcabnro,anrcubmanual" & _
'                                                ",anrrangnro,estrnro,facnro,tenro,Ternro,tipnro" & _
'                                                ",anrcubvalor1,anrcubvalor2) VALUES (" & _
'                                                rsAnrCab!anrcabnro & ",0," & rsRango!anrrangnro & "," & _
'                                                rsEstructura!estrnro & "," & FactorTotalizador & "," & _
'                                                rsEstructura!tenro & "," & tercero & ",1," & _
'                                                cubvalor1 & "," & cubvalor2 & ")"
'                                        Else
'                                            StrSql = "UPDATE anrcubo SET" & _
'                                                " anrcubvalor1 = " & Round(rsTot!anrcubvalor1 + cubvalor1, 2) & _
'                                                " ,anrcubvalor2 = " & rs!anrcubvalor2 + cubvalor2 & _
'                                                " WHERE anrcabnro = " & rsAnrCab!anrcabnro & _
'                                                " AND facnro = " & FactorTotalizador & _
'                                                " AND tenro = " & rsEstructura!tenro & _
'                                                " AND estrnro = " & rsEstructura!estrnro & _
'                                                " AND ternro = " & tercero & _
'                                                " AND anrrangnro = " & rsRango!anrrangnro
'                                        End If
'                                        objConn.Execute StrSql, , adExecuteNoRecords
'                                    End If
                                    
                                End If
                            Else
                                '/* Si no proporciona tomo al 100% y la ultima his_estruc del rango*/
                                porcentaje = 100
                                'If Last_OF(rsFactor!facnro) Or Last_OF(rsEstructura!estrnro) Then
                                If Last_OF_Factor() Or Last_OF_estrnro() Then
                                    If Not Last_OF_tenro() Then
                                        monto_total = 0
                                        cant_total = 0
                                    Else
                                    '/*Busco la ultima his_estr dentro del rango*/
                                        StrSql = "SELECT * FROM his_estructura " & _
                                            " WHERE his_estructura.ternro = " & tercero & _
                                            " AND his_estructura.tenro = " & rsEstructura!tenro & _
                                            " AND his_estructura.htetdesde <= " & ConvFecha(Dia_Fin_Per_Analizado) & _
                                            " AND (his_estructura.htethasta >= " & ConvFecha(Dia_Inicio_Per_Analizado) & _
                                            " OR his_estructura.htethasta IS NULL) "
                                        OpenRecordset StrSql, objRs
                                        objRs.MoveLast
                                        
                                        If Not objRs.EOF Then
                                        
                                            StrSql = "SELECT * FROM anrcubo" & _
                                                " WHERE anrcabnro = " & rsAnrCab!anrcabnro & _
                                                " AND facnro = " & rsHistCon!facnro & _
                                                " AND tenro = " & objRs!tenro & _
                                                " AND estrnro = " & objRs!estrnro & _
                                                " AND ternro = " & tercero & _
                                                " AND anrrangnro = " & rsRango!anrrangnro
                                            OpenRecordset StrSql, rs
                                            
                                            cubvalor1 = monto_total * porcentaje / 100
                                            cubvalor2 = cant_total * porcentaje / 100
                                            
                                            If rs.EOF Then
                                                '/* Creo el cubo */
                                                StrSql = "INSERT INTO anrcubo(anrcabnro,anrcubmanual" & _
                                                    ",anrrangnro,estrnro,facnro,tenro,Ternro,tipnro" & _
                                                    ",anrcubvalor1,anrcubvalor2"
                                                StrSql = StrSql & ") VALUES (" & _
                                                    rsAnrCab!anrcabnro & ",0," & rsRango!anrrangnro & "," & _
                                                    objRs!estrnro & "," & rsHistCon!facnro & "," & _
                                                    objRs!tenro & "," & tercero & ",1" & _
                                                    "," & cubvalor1 & "," & cubvalor2
                                                StrSql = StrSql & ")"
                                            Else
                                                StrSql = "UPDATE anrcubo SET" & _
                                                    " anrcubvalor1 = " & rs!anrcubvalor1 + cubvalor1 & _
                                                    " ,anrcubvalor2 = " & rs!anrcubvalor2 + cubvalor2
                                                StrSql = StrSql & " WHERE anrcabnro = " & rsAnrCab!anrcabnro & _
                                                    " AND facnro = " & rsHistCon!facnro & _
                                                    " AND tenro = " & objRs!tenro & _
                                                    " AND estrnro = " & objRs!estrnro & _
                                                    " AND ternro = " & tercero & _
                                                    " AND anrrangnro = " & rsRango!anrrangnro
                                            End If
                                            objConn.Execute StrSql, , adExecuteNoRecords
                                            
'                                            'FZG 25/07/2003
'                                            'Actualizo Totalizador
'                                            If Totaliza Then
'                                                StrSql = "SELECT * FROM anrcubo" & _
'                                                    " WHERE anrcabnro = " & rsAnrCab!anrcabnro & _
'                                                    " AND facnro = " & FactorTotalizador & _
'                                                    " AND tenro = " & rsEstructura!tenro & _
'                                                    " AND estrnro = " & rsEstructura!estrnro & _
'                                                    " AND ternro = " & tercero & _
'                                                    " AND anrrangnro = " & rsRango!anrrangnro
'                                                OpenRecordset StrSql, rsTot
'
'                                                If rsTot.EOF Then
'                                                    ' Creo el cubo
'                                                    StrSql = "INSERT INTO anrcubo(anrcabnro,anrcubmanual" & _
'                                                        ",anrrangnro,estrnro,facnro,tenro,Ternro,tipnro" & _
'                                                        ",anrcubvalor1,anrcubvalor2) VALUES (" & _
'                                                        rsAnrCab!anrcabnro & ",0," & rsRango!anrrangnro & "," & _
'                                                        rsEstructura!estrnro & "," & FactorTotalizador & "," & _
'                                                        rsEstructura!tenro & "," & tercero & ",1," & _
'                                                        cubvalor1 & "," & cubvalor2 & ")"
'                                                Else
'                                                    StrSql = "UPDATE anrcubo SET" & _
'                                                        " anrcubvalor1 = " & Round(rsTot!anrcubvalor1 + cubvalor1, 2) & _
'                                                        " ,anrcubvalor2 = " & rs!anrcubvalor2 + cubvalor2 & _
'                                                        " WHERE anrcabnro = " & rsAnrCab!anrcabnro & _
'                                                        " AND facnro = " & FactorTotalizador & _
'                                                        " AND tenro = " & rsEstructura!tenro & _
'                                                        " AND estrnro = " & rsEstructura!estrnro & _
'                                                        " AND ternro = " & tercero & _
'                                                        " AND anrrangnro = " & rsRango!anrrangnro
'                                                End If
'                                                objConn.Execute StrSql, , adExecuteNoRecords
'                                            End If
                                            
                                            monto_total = 0
                                            cant_total = 0
                                            
                                            
                                        End If
                                        objRs.Close
                                        
                                    End If
                                End If
                                
                            End If
                        
                            rsEstructura.MoveNext
                        Loop
                
                rsHistCon.MoveNext
            Loop
    
        rsHistliq.MoveNext
    Loop
    
siguientelegajo:
        Progreso = Progreso + IncPorc
       ' Actualizo el progreso
       StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & " WHERE bpronro = " & NroProceso
       objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
        rsFiltro.MoveNext
    Loop
        
        MesActual = MesActual + 1
    Loop 'MesActual
    
    MesActual = 1
    AnioActual = AnioActual + 1
Loop 'AnioActual

    rsRango.MoveNext
Loop
SumPorcTiempo = CDbl(SumPorcTiempo) + CDbl(PorcTiempo)
End Sub


Private Sub ConceptosNexus_Nuevo_NEXUS_RHPRO(Nro_Analisis As Long, Filtrar As Boolean)

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

'Variables para los first y last
Dim PrimerFactOri As Boolean
Dim TipoEstr As Long
Dim EstrAct As Long
Dim FactOri As Long
Dim MiConcepto As String

Dim rs As New ADODB.Recordset

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


'Código -------------------------------------------------------------------

'Abro la conexion para Nexus
OpenConnection strConexionNexus, objConnNexus
'OpenConnection strConexionNexusConf, objConnNexusConf

' --- fgz 07/07/2003---------
'Obtengo la cabecera y el filtro
Call ObtenerCabecerayFiltro(Nro_Analisis, rsAnrCab, Filtrar, rsFiltro, cant_flt, Ok)

If Not Ok Then
    Exit Sub
End If

'Estoy reprocesando
Call PurgarCubo(Nro_Analisis, 7)
' ---------------------------

StrSql = "SELECT * FROM anrrangofec" & _
         " WHERE anrrangofec.anrcabnro = " & rsAnrCab!anrcabnro
                        
OpenRecordset StrSql, rsRango
                    
Do While Not rsRango.EOF
    ' separo el procesamiento en cada uno de los rangos definidos
    ' Recupero el inicio y fin del periodo a analizar
    
    Fin_Per_Analizado = rsRango!anrrangfechasta
    Inicio_Per_Analizado = rsRango!anrrangfecdesde

' divido en periodos por mes
    MesInicio = Month(Inicio_Per_Analizado)
    MesFin = Month(Fin_Per_Analizado)
    AnioInicio = Year(Inicio_Per_Analizado)
    AnioFin = Year(Fin_Per_Analizado)

    MesActual = MesInicio
    AnioActual = AnioInicio
    
Do While AnioActual <= AnioFin

    Do While (MesActual <= 12 And AnioActual < AnioFin) Or (MesActual <= MesFin And AnioActual <= AnioFin)
        ' dia de inicio del periodo a analizar
        If MesActual < 10 Then
            Dia_Inicio_Per_Analizado = CDate("01/0" & MesActual & "/" & AnioActual)
        Else
            Dia_Inicio_Per_Analizado = CDate("01/" & MesActual & "/" & AnioActual)
        End If
        
        ' dia de fin del periodo a analizar
        ' Ultimo dia del mes actual
        If MesActual <> 12 Then
            AuxDia = Day(CDate("01/" & MesActual + 1 & "/" & AnioActual) - 1)
        Else
            AuxDia = 31
        End If
        
        If MesActual < 10 Then
            Dia_Fin_Per_Analizado = CDate(AuxDia & "/0" & MesActual & "/" & AnioActual)
        Else
            Dia_Fin_Per_Analizado = CDate(AuxDia & "/" & MesActual & "/" & AnioActual)
        End If

        ' obtengo el conjunto de legajos a procesar
        Call ObtenerLegajos(1, Filtrar, rsAnrCab!anrcabnro, rsFiltro, Dia_Inicio_Per_Analizado, Dia_Fin_Per_Analizado)
    
        Progreso = 1
        If Not rsFiltro.EOF Then
            IncPorc = 100 / rsFiltro.RecordCount
        End If
    
        ' para probar
       StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & " WHERE bpronro = " & NroProceso
       objConn.Execute StrSql, , adExecuteNoRecords
    
    
    
        perpago_desde = Year(Dia_Inicio_Per_Analizado)
        If (Month(Dia_Inicio_Per_Analizado) < 10) Then
          perpago_desde = perpago_desde & "0"
        End If
        perpago_desde = perpago_desde & Month(Dia_Inicio_Per_Analizado)
        
        perpago_hasta = Year(Dia_Fin_Per_Analizado)
        If (Month(Dia_Fin_Per_Analizado) < 10) Then
          perpago_hasta = perpago_hasta & "0"
        End If
        perpago_hasta = perpago_hasta & Month(Dia_Fin_Per_Analizado)

' Segun el intervalo de análisis, se determinan las cadenas en formato AAAAMM que son utilizadas
' para acceder a la tabla HISTLIQ. O.D.A. 30/06/2003
    
    Do While Not rsFiltro.EOF
      If (cant_flt > 0) And (rsFiltro!cant_te < cant_flt) Then
          GoTo siguientelegajo
      End If
' Cuando hay establecido un filtro, se debe verificar que el empleado verifique
' todos los filtros en el intervalo de tiempo analizado. El control se hace de
' esta forma, para considerar en forma correcta los casos en donde existe más de
' un registro en HIS_ESTRUCTURA, para el mismo empleado y para el mismo tipo de
' estructura, que satisfacen el intervalo de tiempo.
' Legajo 387589, tenro 36, Maciel, para Mayo 2003 O.D.A. 04/07/2003


' Recorre para el analisis las tablas de nexus segun los factores configurados
' Comienzo con las tablas de Nexus
        StrSql = " SELECT legaliq.periodo_pago, legaliq.nro_corr_liq, legaliq.nro_leg FROM histliq " & _
                 " INNER JOIN legaliq " & _
                 " ON    legaliq.periodo_pago   = histliq.periodo_pago" & _
                 " AND   legaliq.nro_corr_liq   = histliq.nro_corr_liq" & _
                 " WHERE histliq.liq_confirmada = 'S' " & _
                 " AND   histliq.periodo_pago   = '" & perpago_desde & "'" & _
                 " AND   legaliq.nro_leg        = " & rsFiltro!empleg & _
                 " GROUP BY legaliq.periodo_pago, legaliq.nro_corr_liq, legaliq.nro_leg " & _
                 " ORDER BY nro_leg"
                 
        ' FGZ - 15/08/2003
        ' si es legajo confidencial ==> utilizo otra conexion
        If UCase(rsFiltro!estrcodext) = "ROCA" Then
            OpenRecordsetNexusConf StrSql, rsHistliq
            'OpenRecordsetNexus StrSql, rsHistliq
        Else
            OpenRecordsetNexus StrSql, rsHistliq
        End If
        
' Con esta forma de trabajo, se está accediendo a los procesos de liquidacion de un periodo
' en particular, por lo que no hay necesidad de tener un intervalo de periodos. O.D.A. 04/07/2003
        
        'FGZ - 11/09/2003
        NombreBD = "RHPRO"
        
        Do While Not rsHistliq.EOF
        
            StrSql = "SELECT conc.conccod, conc.concabr, afact.*, afact_ori.*, acab_fact.*, hcon.* " & _
                " FROM " & NombreBD & ":anrcab_fact acab_fact " & _
                " INNER JOIN " & NombreBD & ":anrfact_ori afact_ori ON afact_ori.facnro = acab_fact.facnro " & _
                " AND afact_ori.tipfacnro = 7" & _
                " INNER JOIN " & NombreBD & ":anrfactor afact ON afact.facnro = acab_fact.facnro" & _
                " INNER JOIN " & NombreBD & ":concepto conc ON conc.concnro = afact_ori.faccodorig" & _
                " INNER JOIN histcon hcon ON hcon.periodo_pago = " & rsHistliq!periodo_pago & _
                " AND   hcon.nro_corr_liq = " & rsHistliq!nro_corr_liq & _
                " AND   hcon.nro_leg      = " & rsHistliq!nro_leg & _
                " WHERE acab_fact.anrcabnro = " & rsAnrCab!anrcabnro & _
                " AND   hcon.estr_liq         = conc.conccod[1,8]" & _
                " AND   hcon.cod_cpto         = conc.conccod[9,12]" & _
                " ORDER BY hcon.nro_leg, afact_ori.facnro"
                
                If UCase(rsFiltro!estrcodext) = "ROCA" Then
                    OpenRecordsetNexusConf StrSql, rsHistCon
                    'OpenRecordsetNexus StrSql, rsHistliq
                Else
                    OpenRecordsetNexus StrSql, rsHistCon
                End If
        
            If Not rsHistCon.EOF Then
                'Para el simular el first_of
                PrimerFactOri = True
                'Para el simular el last_of en la tabla anrfact_ori
                FactOri = rsHistCon!facnro
            End If
            
            Do While Not rsHistCon.EOF
            
                '/* Busco si es factor totalizador
                StrSql = "SELECT * FROM anrfact_tot" & _
                        " WHERE facnro = " & rsHistCon!facnro
                OpenRecordset StrSql, rsFactorTotalizador
                    
                If Not rsFactorTotalizador.EOF Then
                    Totaliza = True
                    ' codigo de factor con el cual se inserta en el cubo
                    FactorTotalizador = rsFactorTotalizador!facnrotot
                Else
                    Totaliza = False
                    FactorTotalizador = 0
                End If
            
                
'                Do While Not rsHistCon.EOF
                    
                    tercero = rsFiltro!Ternro
                    
                    ' fgz 30/05/2003
                    ' Aca deberia controlar que los de la cabecera que viene como parametro
                    ' tengan el campo "anrrangorepro" esten en TRUE
                    ' " WHERE anrrangofec.anrrangorepro = -1 AND anrrangofec.anrcabnro = " & rsAnrcab!anrcabnro &
                    ' Este WHERE estaba en el SELECT de abajo, pero la columna ANRRANGOREPRO no existe!
                    ' O.D.A. 09/06/2003
                                                            
                    ' FGZ ----------------------------------------
                    ' esto estaba funcionando hasta el 10/07/2003
                    '
                    'StrSql = "SELECT * FROM anrrangofec" & _
                    '    " WHERE anrrangofec.anrcabnro = " & rsAnrCab!anrcabnro & _
                    '    " AND anrrangofec.anrrangfecdesde <= " & ConvFecha(rsHistliq!fec_liq) & _
                    '    " AND anrrangofec.anrrangfechasta >= " & ConvFecha(rsHistliq!fec_liq)
                    '
                    'OpenRecordset StrSql, rsRango
                    
                    'Do While Not rsRango.EOF
                    
                    '    StrSql = "SELECT * FROM his_estructura" & _
                    '        " WHERE his_estructura.ternro = " & tercero & _
                    '        " AND his_estructura.htetdesde <= " & ConvFecha(rsRango!anrrangfechasta) & _
                    '        " AND (his_estructura.htethasta >= " & ConvFecha(rsRango!anrrangfecdesde) & _
                    '        " OR his_estructura.htethasta IS NULL)" & _
                    '        " ORDER BY ternro,tenro,estrnro"
                    ' FGZ ----------------------------------------
             
                        ' FGZ 10/07/2003--------------------------
                        Call ObtenerEstructuras(Filtrar, tercero, Dia_Inicio_Per_Analizado, Dia_Fin_Per_Analizado, rsEstructura)
                        
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
                            monto_total = monto_total + rsHistCon!importe_final
                            cant_total = cant_total + (0 & rsHistCon!Cantidad)
                            
                            '/* Calculo los dias de rango entre las fechas del rango y
                            ' el his_estruct para proporcionar*/
                            If rsHistCon!facpropor = -1 Then
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
                                
                                If Last_OF_Factor() Or Last_OF_estrnro() Then
                                    
                                    cubvalor1 = monto_total * porcentaje / 100
                                    cubvalor2 = cant_total * porcentaje / 100
                                    
                                    StrSql = "SELECT * FROM anrcubo" & _
                                        " WHERE anrcabnro = " & rsAnrCab!anrcabnro & _
                                        " AND facnro = " & rsHistCon!facnro & _
                                        " AND tenro = " & rsEstructura!tenro & _
                                        " AND estrnro = " & rsEstructura!estrnro & _
                                        " AND ternro = " & tercero & _
                                        " AND anrrangnro = " & rsRango!anrrangnro
                                    OpenRecordset StrSql, rs
        
                                    'Si el cubo no existe lo creo
                                    If rs.EOF Then
                                    '/* Creo el cubo */
                                        StrSql = "INSERT INTO anrcubo(anrcabnro,anrcubmanual" & _
                                            ",anrrangnro,estrnro,facnro,tenro,Ternro,tipnro" & _
                                            ",anrcubvalor1,anrcubvalor2) VALUES (" & _
                                            rsAnrCab!anrcabnro & ",0," & rsRango!anrrangnro & "," & _
                                            rsEstructura!estrnro & "," & rsHistCon!facnro & "," & _
                                            rsEstructura!tenro & "," & tercero & ",1"
                                    End If
                                    
                                    monto_saldo = (monto_total - cubvalor1 - monto_saldo)
                                    cant_saldo = (cant_total - cubvalor2 - cant_saldo)
                                    monto_total = 0
                                    cant_total = 0
                                    
                                    '* Para que no quede saldo cuando proporciona */
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
                                            " ,anrcubvalor2 = " & rs!anrcubvalor2 + cubvalor2 & _
                                            " WHERE anrcabnro = " & rsAnrCab!anrcabnro & _
                                            " AND facnro = " & rsHistCon!facnro & _
                                            " AND tenro = " & rsEstructura!tenro & _
                                            " AND estrnro = " & rsEstructura!estrnro & _
                                            " AND ternro = " & tercero & _
                                            " AND anrrangnro = " & rsRango!anrrangnro
                                    Else
                                        StrSql = StrSql & "," & cubvalor1 & "," & cubvalor2 & ")"
                                    End If
                                    
                                    objConn.Execute StrSql, , adExecuteNoRecords
                                    
                                    'FZG 25/07/2003
                                    'Actualizo Totalizador
                                    If Totaliza Then
                                        StrSql = "SELECT * FROM anrcubo" & _
                                            " WHERE anrcabnro = " & rsAnrCab!anrcabnro & _
                                            " AND facnro = " & FactorTotalizador & _
                                            " AND tenro = " & rsEstructura!tenro & _
                                            " AND estrnro = " & rsEstructura!estrnro & _
                                            " AND ternro = " & tercero & _
                                            " AND anrrangnro = " & rsRango!anrrangnro
                                        OpenRecordset StrSql, rsTot
                                        
                                        If rsTot.EOF Then
                                            ' Creo el cubo
                                            StrSql = "INSERT INTO anrcubo(anrcabnro,anrcubmanual" & _
                                                ",anrrangnro,estrnro,facnro,tenro,Ternro,tipnro" & _
                                                ",anrcubvalor1,anrcubvalor2) VALUES (" & _
                                                rsAnrCab!anrcabnro & ",0," & rsRango!anrrangnro & "," & _
                                                rsEstructura!estrnro & "," & FactorTotalizador & "," & _
                                                rsEstructura!tenro & "," & tercero & ",1," & _
                                                cubvalor1 & "," & cubvalor2 & ")"
                                        Else
                                            StrSql = "UPDATE anrcubo SET" & _
                                                " anrcubvalor1 = " & Round(rsTot!anrcubvalor1 + cubvalor1, 2) & _
                                                " ,anrcubvalor2 = " & rs!anrcubvalor2 + cubvalor2 & _
                                                " WHERE anrcabnro = " & rsAnrCab!anrcabnro & _
                                                " AND facnro = " & FactorTotalizador & _
                                                " AND tenro = " & rsEstructura!tenro & _
                                                " AND estrnro = " & rsEstructura!estrnro & _
                                                " AND ternro = " & tercero & _
                                                " AND anrrangnro = " & rsRango!anrrangnro
                                        End If
                                        objConn.Execute StrSql, , adExecuteNoRecords
                                    End If
                                    
                                End If
                            Else
                                '/* Si no proporciona tomo al 100% y la ultima his_estruc del rango*/
                                porcentaje = 100
                                'If Last_OF(rsFactor!facnro) Or Last_OF(rsEstructura!estrnro) Then
                                If Last_OF_Factor() Or Last_OF_estrnro() Then
                                    If Not Last_OF_tenro() Then
                                        monto_total = 0
                                        cant_total = 0
                                    Else
                                    '/*Busco la ultima his_estr dentro del rango*/
                                        StrSql = "SELECT * FROM his_estructura " & _
                                            " WHERE his_estructura.ternro = " & tercero & _
                                            " AND his_estructura.tenro = " & rsEstructura!tenro & _
                                            " AND his_estructura.htetdesde <= " & ConvFecha(Dia_Fin_Per_Analizado) & _
                                            " AND (his_estructura.htethasta >= " & ConvFecha(Dia_Inicio_Per_Analizado) & _
                                            " OR his_estructura.htethasta IS NULL) "
                                        OpenRecordset StrSql, objRs
                                        objRs.MoveLast
                                        
                                        If Not objRs.EOF Then
                                        
                                            StrSql = "SELECT * FROM anrcubo" & _
                                                " WHERE anrcabnro = " & rsAnrCab!anrcabnro & _
                                                " AND facnro = " & rsHistCon!facnro & _
                                                " AND tenro = " & objRs!tenro & _
                                                " AND estrnro = " & objRs!estrnro & _
                                                " AND ternro = " & tercero & _
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
                                                    objRs!estrnro & "," & rsHistCon!facnro & "," & _
                                                    objRs!tenro & "," & tercero & ",1" & _
                                                    "," & cubvalor1 & "," & cubvalor2 & ")"
                                            Else
                                                StrSql = "UPDATE anrcubo SET" & _
                                                    " anrcubvalor1 = " & rs!anrcubvalor1 + cubvalor1 & _
                                                    " ,anrcubvalor2 = " & rs!anrcubvalor2 + cubvalor2 & _
                                                    " WHERE anrcabnro = " & rsAnrCab!anrcabnro & _
                                                    " AND facnro = " & rsHistCon!facnro & _
                                                    " AND tenro = " & objRs!tenro & _
                                                    " AND estrnro = " & objRs!estrnro & _
                                                    " AND ternro = " & tercero & _
                                                    " AND anrrangnro = " & rsRango!anrrangnro
                                            End If
                                            objConn.Execute StrSql, , adExecuteNoRecords
                                            
                                            'FZG 25/07/2003
                                            'Actualizo Totalizador
                                            If Totaliza Then
                                                StrSql = "SELECT * FROM anrcubo" & _
                                                    " WHERE anrcabnro = " & rsAnrCab!anrcabnro & _
                                                    " AND facnro = " & FactorTotalizador & _
                                                    " AND tenro = " & rsEstructura!tenro & _
                                                    " AND estrnro = " & rsEstructura!estrnro & _
                                                    " AND ternro = " & tercero & _
                                                    " AND anrrangnro = " & rsRango!anrrangnro
                                                OpenRecordset StrSql, rsTot
                                                
                                                If rsTot.EOF Then
                                                    ' Creo el cubo
                                                    StrSql = "INSERT INTO anrcubo(anrcabnro,anrcubmanual" & _
                                                        ",anrrangnro,estrnro,facnro,tenro,Ternro,tipnro" & _
                                                        ",anrcubvalor1,anrcubvalor2) VALUES (" & _
                                                        rsAnrCab!anrcabnro & ",0," & rsRango!anrrangnro & "," & _
                                                        rsEstructura!estrnro & "," & FactorTotalizador & "," & _
                                                        rsEstructura!tenro & "," & tercero & ",1," & _
                                                        cubvalor1 & "," & cubvalor2 & ")"
                                                Else
                                                    StrSql = "UPDATE anrcubo SET" & _
                                                        " anrcubvalor1 = " & Round(rsTot!anrcubvalor1 + cubvalor1, 2) & _
                                                        " ,anrcubvalor2 = " & rs!anrcubvalor2 + cubvalor2 & _
                                                        " WHERE anrcabnro = " & rsAnrCab!anrcabnro & _
                                                        " AND facnro = " & FactorTotalizador & _
                                                        " AND tenro = " & rsEstructura!tenro & _
                                                        " AND estrnro = " & rsEstructura!estrnro & _
                                                        " AND ternro = " & tercero & _
                                                        " AND anrrangnro = " & rsRango!anrrangnro
                                                End If
                                                objConn.Execute StrSql, , adExecuteNoRecords
                                            End If
                                            
                                            monto_total = 0
                                            cant_total = 0
                                            
                                            
                                        End If
                                        objRs.Close
                                        
                                    End If
                                End If
                                
                            End If
                        
                            rsEstructura.MoveNext
                        Loop
                
                rsHistCon.MoveNext
            Loop
    
'            rsFactor.MoveNext
'        Loop
    
        rsHistliq.MoveNext
    Loop
    
siguientelegajo:
        Progreso = Progreso + IncPorc
       ' Actualizo el progreso
       StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & " WHERE bpronro = " & NroProceso
       objConn.Execute StrSql, , adExecuteNoRecords
    
        rsFiltro.MoveNext
    Loop
        
        MesActual = MesActual + 1
    Loop 'MesActual
    
    MesActual = 1
    AnioActual = AnioActual + 1
Loop 'AnioActual

    rsRango.MoveNext
Loop

End Sub



Public Sub OpenRecordsetNexus(strSQLQuery As String, ByRef objRs As ADODB.Recordset, Optional lockType As LockTypeEnum = adLockReadOnly)
    'Abre un recordset con la consulta strSQLQuery
    If objRs.State <> adStateClosed Then
        If objRs.lockType <> adLockReadOnly Then objRs.UpdateBatch
        objRs.Close
    End If
    objRs.Open strSQLQuery, objConnNexus, adOpenDynamic, lockType, adCmdText
End Sub

Public Sub OpenRecordsetNexusConf(strSQLQuery As String, ByRef objRs As ADODB.Recordset, Optional lockType As LockTypeEnum = adLockReadOnly)
    'Abre un recordset con la consulta strSQLQuery
    If objRs.State <> adStateClosed Then
        If objRs.lockType <> adLockReadOnly Then objRs.UpdateBatch
        objRs.Close
    End If
    objRs.Open strSQLQuery, objConnNexusConf, adOpenDynamic, lockType, adCmdText
End Sub

Function Last_OF_Factor() As Boolean
Dim resultado As Boolean
Dim Actual As Long

    Actual = rsHistCon!facnro
    'Trato de obtener el próximo
    rsHistCon.MoveNext
    'Si es vacío entonces tengo al último del grupo
    If rsHistCon.EOF Then
        resultado = True
    Else
        'Si el proximo es distinto del actual entonces el actual es el ultimo
        If rsHistCon!facnro <> Actual Then
            resultado = True
        Else
            resultado = False
        End If
    End If
    
    rsHistCon.MovePrevious
    Last_OF_Factor = resultado
    
End Function


Private Sub ConceptosNexus_OLD(Nro_Analisis As Long, Filtrar As Boolean)

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

'Variables para los first y last
Dim PrimerFactOri As Boolean
Dim TipoEstr As Long
Dim EstrAct As Long
Dim FactOri As Long
Dim MiConcepto As String

Dim rs As New ADODB.Recordset

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

'Código -------------------------------------------------------------------

'Abro la conexion para Nexus
OpenConnection strConexionNexus, objConnNexus
OpenConnection strConexionNexusConf, objConnNexusConf

' --- fgz 07/07/2003---------
'Obtengo la cabecera y el filtro
Call ObtenerCabecerayFiltro(Nro_Analisis, rsAnrCab, Filtrar, rsFiltro, cant_flt, Ok)

If Not Ok Then
    Exit Sub
End If

'Estoy reprocesando
Call PurgarCubo(Nro_Analisis, 7)
' ---------------------------

StrSql = "SELECT * FROM anrrangofec" & _
         " WHERE anrrangofec.anrcabnro = " & rsAnrCab!anrcabnro
                        
OpenRecordset StrSql, rsRango
                    
Do While Not rsRango.EOF
    ' separo el procesamiento en cada uno de los rangos definidos
    ' Recupero el inicio y fin del periodo a analizar
    
    Fin_Per_Analizado = rsRango!anrrangfechasta
    Inicio_Per_Analizado = rsRango!anrrangfecdesde

' divido en periodos por mes
    MesInicio = Month(Inicio_Per_Analizado)
    MesFin = Month(Fin_Per_Analizado)
    AnioInicio = Year(Inicio_Per_Analizado)
    AnioFin = Year(Fin_Per_Analizado)

    MesActual = MesInicio
    AnioActual = AnioInicio
    
Do While AnioActual <= AnioFin

    Do While (MesActual <= 12 And AnioActual < AnioFin) Or (MesActual <= MesFin And AnioActual <= AnioFin)
        ' dia de inicio del periodo a analizar
        If MesActual < 10 Then
            Dia_Inicio_Per_Analizado = CDate("01/0" & MesActual & "/" & AnioActual)
        Else
            Dia_Inicio_Per_Analizado = CDate("01/" & MesActual & "/" & AnioActual)
        End If
        
        ' dia de fin del periodo a analizar
        ' Ultimo dia del mes actual
        If MesActual <> 12 Then
            AuxDia = Day(CDate("01/" & MesActual + 1 & "/" & AnioActual) - 1)
        Else
            AuxDia = 31
        End If
        
        If MesActual < 10 Then
            Dia_Fin_Per_Analizado = CDate(AuxDia & "/0" & MesActual & "/" & AnioActual)
        Else
            Dia_Fin_Per_Analizado = CDate(AuxDia & "/" & MesActual & "/" & AnioActual)
        End If

        ' obtengo el conjunto de legajos a procesar
        Call ObtenerLegajos(1, Filtrar, rsAnrCab!anrcabnro, rsFiltro, Dia_Inicio_Per_Analizado, Dia_Fin_Per_Analizado)
    
        Progreso = 0
        If Not rsFiltro.EOF Then
            IncPorc = 100 / rsFiltro.RecordCount
        End If
    
        perpago_desde = Year(Dia_Inicio_Per_Analizado)
        If (Month(Dia_Inicio_Per_Analizado) < 10) Then
          perpago_desde = perpago_desde & "0"
        End If
        perpago_desde = perpago_desde & Month(Dia_Inicio_Per_Analizado)
        
        perpago_hasta = Year(Dia_Fin_Per_Analizado)
        If (Month(Dia_Fin_Per_Analizado) < 10) Then
          perpago_hasta = perpago_hasta & "0"
        End If
        perpago_hasta = perpago_hasta & Month(Dia_Fin_Per_Analizado)

' Segun el intervalo de análisis, se determinan las cadenas en formato AAAAMM que son utilizadas
' para acceder a la tabla HISTLIQ. O.D.A. 30/06/2003
    
    Do While Not rsFiltro.EOF
      If (cant_flt > 0) And (rsFiltro!cant_te < cant_flt) Then
          GoTo siguientelegajo
      End If
' Cuando hay establecido un filtro, se debe verificar que el empleado verifique
' todos los filtros en el intervalo de tiempo analizado. El control se hace de
' esta forma, para considerar en forma correcta los casos en donde existe más de
' un registro en HIS_ESTRUCTURA, para el mismo empleado y para el mismo tipo de
' estructura, que satisfacen el intervalo de tiempo.
' Legajo 387589, tenro 36, Maciel, para Mayo 2003 O.D.A. 04/07/2003


' Recorre para el analisis las tablas de nexus segun los factores configurados
' Comienzo con las tablas de Nexus
        StrSql = " SELECT legaliq.periodo_pago, legaliq.nro_corr_liq, legaliq.nro_leg FROM histliq " & _
                 " INNER JOIN legaliq " & _
                 " ON    legaliq.periodo_pago   = histliq.periodo_pago" & _
                 " AND   legaliq.nro_corr_liq   = histliq.nro_corr_liq" & _
                 " WHERE histliq.liq_confirmada = 'S' " & _
                 " AND   histliq.periodo_pago   = '" & perpago_desde & "'" & _
                 " AND   legaliq.nro_leg        = " & rsFiltro!empleg & _
                 " GROUP BY legaliq.periodo_pago, legaliq.nro_corr_liq, legaliq.nro_leg " & _
                 " ORDER BY nro_leg"
                 
        ' FGZ - 15/08/2003
        ' si es legajo confidencial ==> utilizo otra conexion
        If UCase(rsFiltro!estrcodext) = "ROCA" Then
            OpenRecordsetNexusConf StrSql, rsHistliq
        Else
            OpenRecordsetNexus StrSql, rsHistliq
        End If
        
' Con esta forma de trabajo, se está accediendo a los procesos de liquidacion de un periodo
' en particular, por lo que no hay necesidad de tener un intervalo de periodos. O.D.A. 04/07/2003
       
        Do While Not rsHistliq.EOF
            StrSql = "SELECT * FROM anrcab_fact" & _
                " INNER JOIN anrfact_ori ON anrfact_ori.facnro = anrcab_fact.facnro " & _
                " AND anrfact_ori.tipfacnro = 7" & _
                " INNER JOIN anrfactor ON anrfactor.facnro = anrcab_fact.facnro" & _
                " INNER JOIN concepto ON concepto.concnro = anrfact_ori.faccodorig" & _
                " WHERE anrcab_fact.anrcabnro = " & rsAnrCab!anrcabnro & _
                " ORDER BY anrfact_ori.facnro"
        
            OpenRecordset StrSql, rsFactor
            
            If Not rsFactor.EOF Then
                'Para el simular el first_of
                PrimerFactOri = True
                'Para el simular el last_of en la tabla anrfact_ori
                FactOri = rsFactor!facnro
            End If
            
            'Corto para poder enganchar con Nexus
            Do While Not rsFactor.EOF
            
                '/* Busco si es factor totalizador
                StrSql = "SELECT * FROM anrfact_tot" & _
                        " WHERE facnro = " & rsFactor!facnro
                OpenRecordset StrSql, rsFactorTotalizador
                    
                If Not rsFactorTotalizador.EOF Then
                    Totaliza = True
                    ' codigo de factor con el cual se inserta en el cubo
                    FactorTotalizador = rsFactorTotalizador!facnrotot
                Else
                    Totaliza = False
                    FactorTotalizador = 0
                End If
            
                estr_liqNex = Mid(rsFactor!conccod, 1, Len(rsFactor!conccod) - 4)
                cod_cptoNex = Mid(rsFactor!conccod, Len(rsFactor!conccod) - 3, 4)
            
                 StrSql = " SELECT * " & _
                " FROM histcon " & _
                " WHERE histcon.periodo_pago = '" & rsHistliq!periodo_pago & "'" & _
                " AND   histcon.nro_corr_liq = " & rsHistliq!nro_corr_liq & _
                " AND   histcon.nro_leg      = " & rsHistliq!nro_leg & _
                " AND   histcon.estr_liq     = '" & estr_liqNex & "'" & _
                " AND   histcon.cod_cpto     = '" & cod_cptoNex & "'" & _
                " ORDER BY nro_leg"
                
                ' FGZ - 15/08/2003
                ' si es legajo confidencial ==> utilizo otra conexion
                If UCase(rsFiltro!estrcodext) = "ROCA" Then
                    OpenRecordsetNexusConf StrSql, rsHistCon
                Else
                    OpenRecordsetNexus StrSql, rsHistCon
                End If
                
                Do While Not rsHistCon.EOF
                    
                    tercero = rsFiltro!Ternro
                    
                    ' fgz 30/05/2003
                    ' Aca deberia controlar que los de la cabecera que viene como parametro
                    ' tengan el campo "anrrangorepro" esten en TRUE
                    ' " WHERE anrrangofec.anrrangorepro = -1 AND anrrangofec.anrcabnro = " & rsAnrcab!anrcabnro &
                    ' Este WHERE estaba en el SELECT de abajo, pero la columna ANRRANGOREPRO no existe!
                    ' O.D.A. 09/06/2003
                                                            
                    ' FGZ ----------------------------------------
                    ' esto estaba funcionando hasta el 10/07/2003
                    '
                    'StrSql = "SELECT * FROM anrrangofec" & _
                    '    " WHERE anrrangofec.anrcabnro = " & rsAnrCab!anrcabnro & _
                    '    " AND anrrangofec.anrrangfecdesde <= " & ConvFecha(rsHistliq!fec_liq) & _
                    '    " AND anrrangofec.anrrangfechasta >= " & ConvFecha(rsHistliq!fec_liq)
                    '
                    'OpenRecordset StrSql, rsRango
                    
                    'Do While Not rsRango.EOF
                    
                    '    StrSql = "SELECT * FROM his_estructura" & _
                    '        " WHERE his_estructura.ternro = " & tercero & _
                    '        " AND his_estructura.htetdesde <= " & ConvFecha(rsRango!anrrangfechasta) & _
                    '        " AND (his_estructura.htethasta >= " & ConvFecha(rsRango!anrrangfecdesde) & _
                    '        " OR his_estructura.htethasta IS NULL)" & _
                    '        " ORDER BY ternro,tenro,estrnro"
                    ' FGZ ----------------------------------------
             
                        ' FGZ 10/07/2003--------------------------
                        Call ObtenerEstructuras(Filtrar, tercero, Dia_Inicio_Per_Analizado, Dia_Fin_Per_Analizado, rsEstructura)
                        
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
                            monto_total = monto_total + rsHistCon!importe_final
                            cant_total = cant_total + (0 & rsHistCon!Cantidad)
                            
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
                                
                                If Last_OF_Factor() Or Last_OF_estrnro() Then
                                    
                                    cubvalor1 = monto_total * porcentaje / 100
                                    cubvalor2 = cant_total * porcentaje / 100
                                    
                                    StrSql = "SELECT * FROM anrcubo" & _
                                        " WHERE anrcabnro = " & rsAnrCab!anrcabnro & _
                                        " AND facnro = " & rsFactor!facnro & _
                                        " AND tenro = " & rsEstructura!tenro & _
                                        " AND estrnro = " & rsEstructura!estrnro & _
                                        " AND ternro = " & tercero & _
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
                                            rsEstructura!tenro & "," & tercero & ",1"
                                    End If
                                    
                                    monto_saldo = (monto_total - cubvalor1 - monto_saldo)
                                    cant_saldo = (cant_total - cubvalor2 - cant_saldo)
                                    monto_total = 0
                                    cant_total = 0
                                    
                                    '* Para que no quede saldo cuando proporciona */
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
                                            " ,anrcubvalor2 = " & rs!anrcubvalor2 + cubvalor2 & _
                                            " WHERE anrcabnro = " & rsAnrCab!anrcabnro & _
                                            " AND facnro = " & rsFactor!facnro & _
                                            " AND tenro = " & rsEstructura!tenro & _
                                            " AND estrnro = " & rsEstructura!estrnro & _
                                            " AND ternro = " & tercero & _
                                            " AND anrrangnro = " & rsRango!anrrangnro
                                    Else
                                        StrSql = StrSql & "," & cubvalor1 & "," & cubvalor2 & ")"
                                    End If
                                    
                                    objConn.Execute StrSql, , adExecuteNoRecords
                                    
                                    'FZG 25/07/2003
                                    'Actualizo Totalizador
                                    If Totaliza Then
                                        StrSql = "SELECT * FROM anrcubo" & _
                                            " WHERE anrcabnro = " & rsAnrCab!anrcabnro & _
                                            " AND facnro = " & FactorTotalizador & _
                                            " AND tenro = " & rsEstructura!tenro & _
                                            " AND estrnro = " & rsEstructura!estrnro & _
                                            " AND ternro = " & tercero & _
                                            " AND anrrangnro = " & rsRango!anrrangnro
                                        OpenRecordset StrSql, rsTot
                                        
                                        If rsTot.EOF Then
                                            ' Creo el cubo
                                            StrSql = "INSERT INTO anrcubo(anrcabnro,anrcubmanual" & _
                                                ",anrrangnro,estrnro,facnro,tenro,Ternro,tipnro" & _
                                                ",anrcubvalor1,anrcubvalor2) VALUES (" & _
                                                rsAnrCab!anrcabnro & ",0," & rsRango!anrrangnro & "," & _
                                                rsEstructura!estrnro & "," & FactorTotalizador & "," & _
                                                rsEstructura!tenro & "," & tercero & ",1," & _
                                                cubvalor1 & "," & cubvalor2 & ")"
                                        Else
                                            StrSql = "UPDATE anrcubo SET" & _
                                                " anrcubvalor1 = " & Round(rsTot!anrcubvalor1 + cubvalor1, 2) & _
                                                " ,anrcubvalor2 = " & rs!anrcubvalor2 + cubvalor2 & _
                                                " WHERE anrcabnro = " & rsAnrCab!anrcabnro & _
                                                " AND facnro = " & FactorTotalizador & _
                                                " AND tenro = " & rsEstructura!tenro & _
                                                " AND estrnro = " & rsEstructura!estrnro & _
                                                " AND ternro = " & tercero & _
                                                " AND anrrangnro = " & rsRango!anrrangnro
                                        End If
                                        objConn.Execute StrSql, , adExecuteNoRecords
                                    End If
                                    
                                End If
                            Else
                                '/* Si no proporciona tomo al 100% y la ultima his_estruc del rango*/
                                porcentaje = 100
                                'If Last_OF(rsFactor!facnro) Or Last_OF(rsEstructura!estrnro) Then
                                If Last_OF_Factor() Or Last_OF_estrnro() Then
                                    If Not Last_OF_tenro() Then
                                        monto_total = 0
                                        cant_total = 0
                                    Else
                                    '/*Busco la ultima his_estr dentro del rango*/
                                        StrSql = "SELECT * FROM his_estructura " & _
                                            " WHERE his_estructura.ternro = " & tercero & _
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
                                                " AND ternro = " & tercero & _
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
                                                    objRs!tenro & "," & tercero & ",1" & _
                                                    "," & cubvalor1 & "," & cubvalor2 & ")"
                                            Else
                                                StrSql = "UPDATE anrcubo SET" & _
                                                    " anrcubvalor1 = " & rs!anrcubvalor1 + cubvalor1 & _
                                                    " ,anrcubvalor2 = " & rs!anrcubvalor2 + cubvalor2 & _
                                                    " WHERE anrcabnro = " & rsAnrCab!anrcabnro & _
                                                    " AND facnro = " & rsFactor!facnro & _
                                                    " AND tenro = " & objRs!tenro & _
                                                    " AND estrnro = " & objRs!estrnro & _
                                                    " AND ternro = " & tercero & _
                                                    " AND anrrangnro = " & rsRango!anrrangnro
                                            End If
                                            objConn.Execute StrSql, , adExecuteNoRecords
                                            
                                            'FZG 25/07/2003
                                            'Actualizo Totalizador
                                            If Totaliza Then
                                                StrSql = "SELECT * FROM anrcubo" & _
                                                    " WHERE anrcabnro = " & rsAnrCab!anrcabnro & _
                                                    " AND facnro = " & FactorTotalizador & _
                                                    " AND tenro = " & rsEstructura!tenro & _
                                                    " AND estrnro = " & rsEstructura!estrnro & _
                                                    " AND ternro = " & tercero & _
                                                    " AND anrrangnro = " & rsRango!anrrangnro
                                                OpenRecordset StrSql, rsTot
                                                
                                                If rsTot.EOF Then
                                                    ' Creo el cubo
                                                    StrSql = "INSERT INTO anrcubo(anrcabnro,anrcubmanual" & _
                                                        ",anrrangnro,estrnro,facnro,tenro,Ternro,tipnro" & _
                                                        ",anrcubvalor1,anrcubvalor2) VALUES (" & _
                                                        rsAnrCab!anrcabnro & ",0," & rsRango!anrrangnro & "," & _
                                                        rsEstructura!estrnro & "," & FactorTotalizador & "," & _
                                                        rsEstructura!tenro & "," & tercero & ",1," & _
                                                        cubvalor1 & "," & cubvalor2 & ")"
                                                Else
                                                    StrSql = "UPDATE anrcubo SET" & _
                                                        " anrcubvalor1 = " & Round(rsTot!anrcubvalor1 + cubvalor1, 2) & _
                                                        " ,anrcubvalor2 = " & rs!anrcubvalor2 + cubvalor2 & _
                                                        " WHERE anrcabnro = " & rsAnrCab!anrcabnro & _
                                                        " AND facnro = " & FactorTotalizador & _
                                                        " AND tenro = " & rsEstructura!tenro & _
                                                        " AND estrnro = " & rsEstructura!estrnro & _
                                                        " AND ternro = " & tercero & _
                                                        " AND anrrangnro = " & rsRango!anrrangnro
                                                End If
                                                objConn.Execute StrSql, , adExecuteNoRecords
                                            End If
                                            
                                            monto_total = 0
                                            cant_total = 0
                                            
                                            
                                        End If
                                        objRs.Close
                                        
                                    End If
                                End If
                                
                            End If
                        
                            rsEstructura.MoveNext
                        Loop
                
                rsHistCon.MoveNext
            Loop
    
            rsFactor.MoveNext
        Loop
    
        rsHistliq.MoveNext
    Loop
    
siguientelegajo:
        Progreso = Progreso + IncPorc
       ' Actualizo el progreso
       StrSql = "UPDATE batch_proceso SET bprcprogreso = " & CInt(Progreso) & " WHERE bpronro = " & NroProceso
       objConn.Execute StrSql, , adExecuteNoRecords
    
        rsFiltro.MoveNext
    Loop
        
        MesActual = MesActual + 1
    Loop 'MesActual
    
    MesActual = 1
    AnioActual = AnioActual + 1
Loop 'AnioActual

    rsRango.MoveNext
Loop

End Sub


