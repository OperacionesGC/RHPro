Attribute VB_Name = "mdlGerencial"
Option Explicit


'Para Sql server
'Global Const strConexionNexus = "DSN=Nexushr-RHPro;database=nexus;uid=sa;pwd="

'Para Informix
Global Const strConexionNexus = "DSN=Nexushr-RHPro"

Dim objConnNexus As New ADODB.Connection

'Variables de ADO
Dim objRs As New ADODB.Recordset
Dim rsEmp As New ADODB.Recordset
Dim rsAnrCab As New ADODB.Recordset
Dim rsHistliq As New ADODB.Recordset
Dim rsFactor As New ADODB.Recordset
Dim rsFactorTotalizador As New ADODB.Recordset
Dim rsHistCon As New ADODB.Recordset
Dim rsEstructura As New ADODB.Recordset
Dim rsRango As New ADODB.Recordset
Dim rsConc As New ADODB.Recordset
Dim rsAcumDiario As New ADODB.Recordset
Dim rsFiltro As New ADODB.Recordset
Dim rsTot As New ADODB.Recordset
Dim IncPorc As Single
Dim CantFactor As Integer
Dim Progreso As Single
Dim NroProceso As Long

Dim FactorTotalizador As Long
Dim Totaliza As Boolean


Public Sub Main()
Dim fechaDesde As Date
Dim fechaHasta As Date
Dim Fecha As Date
Dim objRs As New ADODB.Recordset
Dim objrsEmpleado As New ADODB.Recordset
Dim strCmdLine  As String
Dim nro_analisis As Long
Dim tipo_factor As Integer
Dim Filtrar As Boolean
Dim pos1 As Integer
Dim pos2 As Integer

On Error GoTo ce

    ' carga las configuraciones basicas, formato de fecha, string de conexion,
    ' tipo de BD y ubicacion del archivo de log
    Call CargarConfiguracionesBasicas

    strCmdLine = Command()
    'strCmdLine = "25953"
    If IsNumeric(strCmdLine) Then
        NroProceso = strCmdLine
    Else
        Exit Sub
    End If
        
    OpenConnection strconexion, objConn
    
    StrSql = "SELECT bprcfecdesde,bprcfechasta,bprcparam FROM batch_proceso WHERE bpronro = " & NroProceso
    objRs.Open StrSql, objConn
    
    'Levanto dos parámetros: el primero es número de análisis
                            'el segundo es el tipo de factor a analizar
                            'el tercero es el nro de cabecera
                            'el cuarto es si se usa o no el filtro de estructuras
    If Not IsNull(objRs!bprcparam) Then
        If Len(objRs!bprcparam) >= 1 Then
            pos1 = 1
            pos2 = InStr(pos1, objRs!bprcparam, ",") - 1
            nro_analisis = Mid(objRs!bprcparam, pos1, pos2)
            
            pos1 = pos2 + 2
            pos2 = InStr(pos1, objRs!bprcparam, ",") - 1
            tipo_factor = CInt(Mid(objRs!bprcparam, pos1, pos2 - pos1 + 1))
            
            pos1 = pos2 + 2
            pos2 = Len(objRs!bprcparam)
            Filtrar = CBool(Mid(objRs!bprcparam, pos1, pos2 - pos1 + 1))
            
        End If
    End If
    
    objRs.Close

    'Cambio el estado del proceso a Procesando
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcestado = 'Procesando' WHERE bpronro = " & NroProceso
    objConn.Execute StrSql, , adExecuteNoRecords
        
    'En funcion del tipo de factor ejecuto un procedimiento u otro.
    ' FGZ 21/07/03
    ' Los factores Totalizadores se controlan en cada estos procedimientos.
    
    Select Case tipo_factor
    Case 4
        Call AcumuladoDiario(nro_analisis, Filtrar)
    Case 5
        Call AcumuladoParcial(nro_analisis, Filtrar)
    Case 6
        Call Licencias(nro_analisis, Filtrar)
    Case 7
        Call ConceptosNexus(nro_analisis, Filtrar)
    End Select
    
    StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado' WHERE bpronro = " & NroProceso
    objConn.Execute StrSql, , adExecuteNoRecords
    
        
    If objConn.State = adStateOpen Then objConn.Close

    Exit Sub
ce:
'    MyRollbackTrans
    StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Error' WHERE bpronro = " & NroProceso
    objConn.Execute StrSql, , adExecuteNoRecords
    If objConn.State = adStateOpen Then objConn.Close
End Sub

Public Sub OpenRecordsetNexus(strSQLQuery As String, ByRef objRs As ADODB.Recordset, Optional lockType As LockTypeEnum = adLockReadOnly)
    'Abre un recordset con la consulta strSQLQuery
    If objRs.State <> adStateClosed Then
        If objRs.lockType <> adLockReadOnly Then objRs.UpdateBatch
        objRs.Close
    End If
    objRs.Open strSQLQuery, objConnNexus, adOpenDynamic, lockType, adCmdText
End Sub

Function Last_OF_Factor() As Boolean
Dim resultado As Boolean
Dim Actual As Long

    Actual = rsFactor!facnro
    'Trato de obtener el próximo
    rsFactor.MoveNext
    'Si es vacío entonces tengo al último del grupo
    If rsFactor.EOF Then
        resultado = True
    Else
        'Si el proximo es distinto del actual entonces el actual es el ultimo
        If rsFactor!facnro <> Actual Then
            resultado = True
        Else
            resultado = False
        End If
    End If
    
    rsFactor.MovePrevious
    Last_OF_Factor = resultado
    
End Function

Function Last_OF_tenro() As Boolean
Dim resultado As Boolean
Dim Actual As Long

    Actual = rsEstructura!tenro
    'Trato de obtener el próximo
    rsEstructura.MoveNext
    'Si es vacío entonces tengo al último del grupo
    If rsEstructura.EOF Then
        resultado = True
    Else
        'Si el proximo es distinto del actual entonces el actual es el ultimo
        If rsEstructura!tenro <> Actual Then
            resultado = True
        Else
            resultado = False
        End If
    End If
    
    rsEstructura.MovePrevious
    Last_OF_tenro = resultado
    
End Function

Function Last_OF_estrnro() As Boolean
Dim resultado As Boolean
Dim Actual As Long

    Actual = rsEstructura!estrnro
    'Trato de obtener el próximo
    rsEstructura.MoveNext
    'Si es vacío entonces tengo al último del grupo
    If rsEstructura.EOF Then
        resultado = True
    Else
        'Si el proximo es distinto del actual entonces el actual es el ultimo
        If rsEstructura!estrnro <> Actual Then
            resultado = True
        Else
            resultado = False
        End If
    End If
    
    rsEstructura.MovePrevious
    Last_OF_estrnro = resultado
    
End Function


Private Sub ConceptosNexus(nro_analisis As Long, Filtrar As Boolean)

'Variables locales
Dim cant_flt As Long
Dim desde As Date
Dim hasta As Date
Dim horas As Single
Dim NroCab As Long
Dim perpago_desde As Long
Dim perpago_hasta As Long

Dim Tercero As Long

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

' --- fgz 07/07/2003---------
'Obtengo la cabecera y el filtro
Call ObtenerCabecerayFiltro(nro_analisis, rsAnrCab, Filtrar, rsFiltro, cant_flt, Ok)

If Not Ok Then
    Exit Sub
End If

'Estoy reprocesando
Call PurgarCubo(nro_analisis, 7)
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
          GoTo siguienteLegajo
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
        OpenRecordsetNexus StrSql, rsHistliq
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
                
                OpenRecordsetNexus StrSql, rsHistCon
            
                Do While Not rsHistCon.EOF
                    
                    Tercero = rsFiltro!Ternro
                    
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
                        Call ObtenerEstructuras(Filtrar, Tercero, Dia_Inicio_Per_Analizado, Dia_Fin_Per_Analizado, rsEstructura)
                        
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
                                        " AND ternro = " & Tercero & _
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
                                            rsEstructura!tenro & "," & Tercero & ",1"
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
                                            " AND ternro = " & Tercero & _
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
                                            " AND ternro = " & Tercero & _
                                            " AND anrrangnro = " & rsRango!anrrangnro
                                        OpenRecordset StrSql, rsTot
                                        
                                        If rsTot.EOF Then
                                            ' Creo el cubo
                                            StrSql = "INSERT INTO anrcubo(anrcabnro,anrcubmanual" & _
                                                ",anrrangnro,estrnro,facnro,tenro,Ternro,tipnro" & _
                                                ",anrcubvalor1,anrcubvalor2) VALUES (" & _
                                                rsAnrCab!anrcabnro & ",0," & rsRango!anrrangnro & "," & _
                                                rsEstructura!estrnro & "," & FactorTotalizador & "," & _
                                                rsEstructura!tenro & "," & Tercero & ",1," & _
                                                cubvalor1 & "," & cubvalor2 & ")"
                                        Else
                                            StrSql = "UPDATE anrcubo SET" & _
                                                " anrcubvalor1 = " & Round(rsTot!anrcubvalor1 + cubvalor1, 2) & _
                                                " ,anrcubvalor2 = " & rs!anrcubvalor2 + cubvalor2 & _
                                                " WHERE anrcabnro = " & rsAnrCab!anrcabnro & _
                                                " AND facnro = " & FactorTotalizador & _
                                                " AND tenro = " & rsEstructura!tenro & _
                                                " AND estrnro = " & rsEstructura!estrnro & _
                                                " AND ternro = " & Tercero & _
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
                                            " WHERE his_estructura.ternro = " & Tercero & _
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
                                                " AND ternro = " & Tercero & _
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
                                                    objRs!tenro & "," & Tercero & ",1" & _
                                                    "," & cubvalor1 & "," & cubvalor2 & ")"
                                            Else
                                                StrSql = "UPDATE anrcubo SET" & _
                                                    " anrcubvalor1 = " & rs!anrcubvalor1 + cubvalor1 & _
                                                    " ,anrcubvalor2 = " & rs!anrcubvalor2 + cubvalor2 & _
                                                    " WHERE anrcabnro = " & rsAnrCab!anrcabnro & _
                                                    " AND facnro = " & rsFactor!facnro & _
                                                    " AND tenro = " & objRs!tenro & _
                                                    " AND estrnro = " & objRs!estrnro & _
                                                    " AND ternro = " & Tercero & _
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
                                                    " AND ternro = " & Tercero & _
                                                    " AND anrrangnro = " & rsRango!anrrangnro
                                                OpenRecordset StrSql, rsTot
                                                
                                                If rsTot.EOF Then
                                                    ' Creo el cubo
                                                    StrSql = "INSERT INTO anrcubo(anrcabnro,anrcubmanual" & _
                                                        ",anrrangnro,estrnro,facnro,tenro,Ternro,tipnro" & _
                                                        ",anrcubvalor1,anrcubvalor2) VALUES (" & _
                                                        rsAnrCab!anrcabnro & ",0," & rsRango!anrrangnro & "," & _
                                                        rsEstructura!estrnro & "," & FactorTotalizador & "," & _
                                                        rsEstructura!tenro & "," & Tercero & ",1," & _
                                                        cubvalor1 & "," & cubvalor2 & ")"
                                                Else
                                                    StrSql = "UPDATE anrcubo SET" & _
                                                        " anrcubvalor1 = " & Round(rsTot!anrcubvalor1 + cubvalor1, 2) & _
                                                        " ,anrcubvalor2 = " & rs!anrcubvalor2 + cubvalor2 & _
                                                        " WHERE anrcabnro = " & rsAnrCab!anrcabnro & _
                                                        " AND facnro = " & FactorTotalizador & _
                                                        " AND tenro = " & rsEstructura!tenro & _
                                                        " AND estrnro = " & rsEstructura!estrnro & _
                                                        " AND ternro = " & Tercero & _
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
    
siguienteLegajo:
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

Private Sub ObtenerCabecerayFiltro(ByVal nro_analisis As Long, ByRef rsAnrCab As ADODB.Recordset, ByRef Filtrar As Boolean, ByRef rs As ADODB.Recordset, ByRef Cantidad As Long, ByRef Ok As Boolean)

Ok = True

StrSql = " SELECT anrcabnro,anrcabfecdesde,anrcabfechasta FROM anrcab " & _
    " WHERE anrcabnro = " & nro_analisis
OpenRecordset StrSql, rsAnrCab
    
If rsAnrCab.EOF Then
    Ok = False
End If

Cantidad = 0
If Filtrar Then
    StrSql = " SELECT COUNT( DISTINCT anrcab_filtro.tenro) AS Cant" & _
             " FROM   anrcab_filtro" & _
             " WHERE  anrcab_filtro.anrcabnro = " & rsAnrCab!anrcabnro
    OpenRecordset StrSql, rs

    If rs.EOF Then
        Cantidad = 0
    Else
        Cantidad = rs!cant
    End If

    If (Cantidad <= 0) Then
        Filtrar = False
    End If
End If

End Sub


Private Sub PurgarCubo(ByVal nro_analisis As Long, ByVal TipoFactor As Integer)

StrSql = "DELETE FROM anrcubo " & _
    " WHERE facnro IN " & _
    " (SELECT facnro FROM anrfactor" & _
    " WHERE tipfacnro = " & TipoFactor & ")" & _
    " AND anrcabnro = " & nro_analisis & _
    " AND anrcubmanual = 0"
    objConn.Execute StrSql, , adExecuteNoRecords

End Sub


Private Sub ObtenerLegajos(ByVal TipoGerencial As Integer, ByVal Filtrar As Boolean, ByVal CabNro As Long, ByRef rsFiltro As ADODB.Recordset, ByVal Dia_Inicio_Per_Analizado As Date, ByVal Dia_Fin_Per_Analizado As Date)

Select Case TipoGerencial
Case 1: 'Conceptos Nexus
' obtengo el conjunto de legajos a procesar
    If Filtrar Then
      StrSql = " SELECT     empleado.empleg, empleado.ternro, COUNT( DISTINCT his_estructura.tenro) as cant_te" & _
               " FROM       empleado " & _
               " INNER JOIN his_estructura " & _
               " ON         his_estructura.ternro     = empleado.ternro " & _
               " AND        his_estructura.htetdesde <= " & ConvFecha(Dia_Fin_Per_Analizado) & _
               " AND       (his_estructura.htethasta >= " & ConvFecha(Dia_Inicio_Per_Analizado) & " OR " & _
               "            his_estructura.htethasta IS NULL) " & _
               " INNER JOIN anrcab_filtro " & _
               " ON         anrcab_filtro.estrnro     = his_estructura.estrnro " & _
               " AND        anrcab_filtro.anrcabnro   = " & CabNro & _
               " GROUP BY   empleado.empleg, empleado.ternro" & _
               " ORDER BY   empleado.empleg"

' Se exige que los empleados cumplan con todas las condiciones especificadas. O.D.A. 27/06/2003

    Else
      StrSql = " SELECT     DISTINCT empleado.empleg, empleado.ternro, 0 as cant_te" & _
               " FROM       empleado " & _
               " INNER JOIN his_estructura " & _
               " ON         his_estructura.ternro     = empleado.ternro " & _
               " AND        his_estructura.htetdesde <= " & ConvFecha(Dia_Fin_Per_Analizado) & _
               " AND       (his_estructura.htethasta >= " & ConvFecha(Dia_Inicio_Per_Analizado) & " OR " & _
               "            his_estructura.htethasta IS NULL) "
    End If
    
Case 2: 'Acumulados Diarios
    If Filtrar Then
        StrSql = " SELECT   his_estructura.tenro, his_estructura.estrnro, his_estructura.ternro, his_estructura.htethasta, his_estructura.htetdesde, COUNT( DISTINCT his_estructura.tenro) as cant_te" & _
               " FROM       his_estructura " & _
               " WHERE      his_estructura.htetdesde <= " & ConvFecha(Dia_Fin_Per_Analizado) & _
               " AND       (his_estructura.htethasta >= " & ConvFecha(Dia_Inicio_Per_Analizado) & " OR " & _
               "            his_estructura.htethasta IS NULL) " & _
               " GROUP BY   his_estructura.ternro, his_estructura.tenro, his_estructura.estrnro, his_estructura.htetdesde, his_estructura.htethasta" & _
               " ORDER BY   his_estructura.ternro, his_estructura.tenro, his_estructura.estrnro"
    Else
      StrSql = " SELECT     DISTINCT his_estructura.tenro, his_estructura.estrnro, his_estructura.ternro, his_estructura.htethasta, his_estructura.htetdesde, 0 as cant_te" & _
               " FROM       his_estructura " & _
               " WHERE      his_estructura.htetdesde <= " & ConvFecha(Dia_Fin_Per_Analizado) & _
               " AND       (his_estructura.htethasta >= " & ConvFecha(Dia_Inicio_Per_Analizado) & " OR " & _
               "            his_estructura.htethasta IS NULL) " & _
               " ORDER BY   his_estructura.ternro, his_estructura.tenro, his_estructura.estrnro"
    End If
    
    
Case 3: 'Acumulados Parciales
    If Filtrar Then
        StrSql = " SELECT   his_estructura.tenro, his_estructura.estrnro, his_estructura.ternro, his_estructura.htethasta, his_estructura.htetdesde, COUNT( DISTINCT his_estructura.tenro) as cant_te" & _
               " FROM       his_estructura " & _
               " WHERE      his_estructura.htetdesde <= " & ConvFecha(Dia_Fin_Per_Analizado) & _
               " AND       (his_estructura.htethasta >= " & ConvFecha(Dia_Inicio_Per_Analizado) & " OR " & _
               "            his_estructura.htethasta IS NULL) " & _
               " GROUP BY   his_estructura.ternro, his_estructura.tenro, his_estructura.estrnro, his_estructura.htetdesde, his_estructura.htethasta" & _
               " ORDER BY   his_estructura.ternro, his_estructura.tenro, his_estructura.estrnro"
    Else
      StrSql = " SELECT     DISTINCT his_estructura.tenro, his_estructura.estrnro, his_estructura.ternro, his_estructura.htethasta, his_estructura.htetdesde, 0 as cant_te" & _
               " FROM       his_estructura " & _
               " WHERE      his_estructura.htetdesde <= " & ConvFecha(Dia_Fin_Per_Analizado) & _
               " AND       (his_estructura.htethasta >= " & ConvFecha(Dia_Inicio_Per_Analizado) & " OR " & _
               "            his_estructura.htethasta IS NULL) " & _
               " ORDER BY   his_estructura.ternro, his_estructura.tenro, his_estructura.estrnro"
    End If

Case 4: 'Licencias
    If Filtrar Then
        StrSql = " SELECT   his_estructura.tenro, his_estructura.estrnro, his_estructura.ternro, his_estructura.htethasta, his_estructura.htetdesde, COUNT( DISTINCT his_estructura.tenro) as cant_te" & _
               " FROM       his_estructura " & _
               " WHERE      his_estructura.htetdesde <= " & ConvFecha(Dia_Fin_Per_Analizado) & _
               " AND       (his_estructura.htethasta >= " & ConvFecha(Dia_Inicio_Per_Analizado) & " OR " & _
               "            his_estructura.htethasta IS NULL) " & _
               " GROUP BY   his_estructura.ternro, his_estructura.tenro, his_estructura.estrnro, his_estructura.htetdesde, his_estructura.htethasta" & _
               " ORDER BY   his_estructura.ternro, his_estructura.tenro, his_estructura.estrnro"
    Else
      StrSql = " SELECT     DISTINCT his_estructura.tenro, his_estructura.estrnro, his_estructura.ternro, his_estructura.htethasta, his_estructura.htetdesde, 0 as cant_te" & _
               " FROM       his_estructura " & _
               " WHERE      his_estructura.htetdesde <= " & ConvFecha(Dia_Fin_Per_Analizado) & _
               " AND       (his_estructura.htethasta >= " & ConvFecha(Dia_Inicio_Per_Analizado) & " OR " & _
               "            his_estructura.htethasta IS NULL) " & _
               " ORDER BY   his_estructura.ternro, his_estructura.tenro, his_estructura.estrnro"
    End If

Case 5: 'Totalizadores
' obtengo el conjunto de legajos a procesar
    If Filtrar Then
      StrSql = " SELECT     empleado.empleg, empleado.ternro, COUNT( DISTINCT his_estructura.tenro) as cant_te" & _
               " FROM       empleado " & _
               " INNER JOIN his_estructura " & _
               " ON         his_estructura.ternro     = empleado.ternro " & _
               " AND        his_estructura.htetdesde <= " & ConvFecha(Dia_Fin_Per_Analizado) & _
               " AND       (his_estructura.htethasta >= " & ConvFecha(Dia_Inicio_Per_Analizado) & " OR " & _
               "            his_estructura.htethasta IS NULL) " & _
               " INNER JOIN anrcab_filtro " & _
               " ON         anrcab_filtro.estrnro     = his_estructura.estrnro " & _
               " AND        anrcab_filtro.anrcabnro   = " & CabNro & _
               " GROUP BY   empleado.empleg, empleado.ternro" & _
               " ORDER BY   empleado.empleg"

' Se exige que los empleados cumplan con todas las condiciones especificadas. O.D.A. 27/06/2003

    Else
      StrSql = " SELECT     DISTINCT empleado.empleg, empleado.ternro, 0 as cant_te" & _
               " FROM       empleado " & _
               " INNER JOIN his_estructura " & _
               " ON         his_estructura.ternro     = empleado.ternro " & _
               " AND        his_estructura.htetdesde <= " & ConvFecha(Dia_Fin_Per_Analizado) & _
               " AND       (his_estructura.htethasta >= " & ConvFecha(Dia_Inicio_Per_Analizado) & " OR " & _
               "            his_estructura.htethasta IS NULL) "
    End If

Case Else
End Select
    
OpenRecordset StrSql, rsFiltro

End Sub


Private Sub ObtenerEstructuras(ByVal Filtrar As Boolean, ByVal Tercero As Long, ByVal FechaInicio As Date, ByVal FechaFin As Date, ByRef rs As ADODB.Recordset)

If Filtrar Then
    StrSql = "SELECT * FROM his_estructura" & _
        " WHERE his_estructura.ternro = " & Tercero & _
        " AND his_estructura.htetdesde <= " & ConvFecha(FechaFin) & _
        " AND (his_estructura.htethasta >= " & ConvFecha(FechaInicio) & _
        " OR his_estructura.htethasta IS NULL)" & _
        " ORDER BY ternro,tenro,estrnro"
Else ' no se usa el filtro ==> todas las estructuras
    StrSql = "SELECT * FROM his_estructura" & _
        " WHERE his_estructura.ternro = " & Tercero & _
        " AND his_estructura.htetdesde <= " & ConvFecha(FechaFin) & _
        " AND (his_estructura.htethasta >= " & ConvFecha(FechaInicio) & _
        " OR his_estructura.htethasta IS NULL)" & _
        " ORDER BY ternro,tenro,estrnro"
End If

OpenRecordset StrSql, rs

End Sub


Private Sub AcumuladoDiario(nro_analisis As Long, Filtrar As Boolean)

'Variables locales
Dim desde As Date
Dim hasta As Date
Dim horas As Single
Dim NroCab As Long
Dim Tercero As Long

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
Dim Ok As Boolean
Dim cant_flt As Long

Dim rs As New ADODB.Recordset

'Código -------------------------------------------------------------------

'Abro la conexion para Nexus
OpenConnection strConexionNexus, objConnNexus

'Obtengo la cabecera
Call ObtenerCabecerayFiltro(nro_analisis, rsAnrCab, Filtrar, rsFiltro, cant_flt, Ok)

If Not Ok Then
    Exit Sub
End If

'Estoy reprocesando
Call PurgarCubo(nro_analisis, 4)

'Comienzo el procesamiento

'/* Recorre para el analisis los acumulados diario de tipos de horas configurados */
StrSql = "SELECT * FROM anrcab_fact" & _
    " INNER JOIN anrfact_ori ON anrfact_ori.facnro = anrcab_fact.facnro" & _
    " AND anrfact_ori.tipfacnro = 4" & _
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
    
Progreso = 0
CantFactor = 0
If Not rsFactor.EOF Then
    CantFactor = rsFactor.RecordCount
End If
    
' obtengo el conjunto de legajos a procesar
Call ObtenerLegajos(1, Filtrar, rsAnrCab!anrcabnro, rsFiltro, rsAnrCab!anrcabfecdesde, rsAnrCab!anrcabfechasta)
    
'Recorro los acumulados diarios
Do While Not rsFactor.EOF
    
        'Busco si es factor totalizador
        StrSql = "SELECT * " & _
                " FROM   anrfact_tot, anrcab_fact" & _
                " WHERE  anrfact_tot.facnro = " & rsFactor!facnro & _
                " AND    anrcab_fact.facnro   = anrfact_tot.facnro " & _
                " AND    anrcab_fact.anrcabnro = " & rsAnrCab!anrcabnro
        OpenRecordset StrSql, rsFactorTotalizador
            
            
        'Busco si es factor totalizador
        'StrSql = "SELECT * FROM anrfact_tot" & _
        '        " WHERE facnro = " & rsFactor!facnro
        'OpenRecordset StrSql, rsFactorTotalizador
            
        If Not rsFactorTotalizador.EOF Then
            Totaliza = True
            ' codigo de factor con el cual se inserta en el cubo
            FactorTotalizador = rsFactorTotalizador!facnrotot
        Else
            Totaliza = False
            FactorTotalizador = 0
        End If

    
    
    ' voy nuevamente al primer legajo del filtro
    rsFiltro.MoveFirst
    
    Do While Not rsFiltro.EOF
        If (cant_flt > 0) And (rsFiltro!cant_te < cant_flt) Then
            GoTo siguienteLegajo
        End If
    
        StrSql = " SELECT * " & _
            " FROM gti_acumdiario " & _
            " WHERE adfecha <= " & ConvFecha(rsAnrCab!anrcabfechasta) & _
            " AND adfecha >= " & ConvFecha(rsAnrCab!anrcabfecdesde) & _
            " AND thnro = " & rsFactor!faccodorig & _
            " AND ternro = " & rsFiltro!Ternro & _
            " ORDER BY ternro"
        OpenRecordset StrSql, rsAcumDiario
        
        
        If Not rsAcumDiario.EOF Then
            IncPorc = ((100 / CantFactor) * (100 / rsAcumDiario.RecordCount)) / 100
        End If
    
        Do While Not rsAcumDiario.EOF
                    
                    StrSql = "SELECT * FROM anrrangofec" & _
                        " WHERE anrrangofec.anrcabnro = " & rsAnrCab!anrcabnro & _
                        " AND anrrangofec.anrrangfecdesde <= " & ConvFecha(rsAcumDiario!adfecha) & _
                        " AND anrrangofec.anrrangfechasta >= " & ConvFecha(rsAcumDiario!adfecha)
                    OpenRecordset StrSql, rsRango
                    
                    
                    Do While Not rsRango.EOF
                        
                        ' Obtengo las estructuras
                        Call ObtenerEstructuras(Filtrar, rsAcumDiario!Ternro, rsRango!anrrangfecdesde, rsRango!anrrangfechasta, rsEstructura)
                        
                        'StrSql = "SELECT * FROM his_estructura" & _
                        '    " WHERE his_estructura.ternro = " & rsAcumDiario!Ternro & _
                        '    " AND his_estructura.htetdesde <= " & ConvFecha(rsRango!anrrangfechasta) & _
                        '    " AND (his_estructura.htethasta >= " & ConvFecha(rsRango!anrrangfecdesde) & _
                        '    " OR his_estructura.htethasta IS NULL)" & _
                        '    " ORDER BY ternro,tenro,estrnro"
                        'OpenRecordset StrSql, rsEstructura
                                        
                        If Not rsEstructura.EOF Then
                            TipoEstr = rsEstructura!tenro
                            EstrAct = rsEstructura!estrnro
                        End If
                    
                        Do While Not rsEstructura.EOF
    
                            If PrimerFactOri Then
                                cantdiasper = DateDiff("d", rsRango!anrrangfecdesde, rsRango!anrrangfechasta) + 1
                                monto_total = 0
                                cant_total = 0
                                cant_saldo = 0
                                PrimerFactOri = False
                            End If
                                           
                            '/* Acumulo por Factor */
                            monto_total = monto_total + rsAcumDiario!adcanthoras
                            cant_total = cant_total
                            
                            '/* Calculo los dias de rango entre las fechas del rango y
                            ' el his_estruct para proporcionar*/
                            If rsFactor!facpropor = -1 Then
                                If rsEstructura!htetdesde < rsRango!anrrangfecdesde Then
                                        If rsEstructura!htethasta < rsRango!anrrangfechasta And (Not IsNull(rsEstructura!htethasta)) Then
                                            cantdiasran = DateDiff("d", rsRango!anrrangfecdesde, rsEstructura!htethasta) + 1
                                        Else
                                            cantdiasran = DateDiff("d", rsRango!anrrangfecdesde, rsRango!anrrangfechasta) + 1
                                        End If
                                Else
                                    If (rsEstructura!htethasta < rsRango!anrrangfechasta) And (Not IsNull(rsEstructura!htethasta)) Then
                                        cantdiasran = DateDiff("d", rsEstructura!htetdesde, rsEstructura!htethasta) + 1
                                    Else
                                        cantdiasran = DateDiff("d", rsEstructura!htetdesde, rsRango!anrrangfechasta) + 1
                                    End If
                                End If
                                
                                '/* Porcentaje segun la cant. de dias en la his_estrutura */
                                porcentaje = cantdiasran * 100 / cantdiasper
                                
                                If Last_OF_Factor() Or Last_OF_estrnro() Then
                                    
                                    'cubvalor1 = monto_total * porcentaje / 100
                                    'cubvalor2 = cant_total * porcentaje / 100
                                    
                                    ' FGZ  18/07/2003 ------------------
                                    cubvalor1 = monto_total
                                    cubvalor2 = cant_total
                                    ' ----------------------------------
                                    
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
                                            rsEstructura!tenro & "," & rsAcumDiario!Ternro & ",1"
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
                                            " AND ternro = " & rsAcumDiario!Ternro & _
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
                                            " AND ternro = " & rsAcumDiario!Ternro & _
                                            " AND anrrangnro = " & rsRango!anrrangnro
                                        OpenRecordset StrSql, rsTot
                                        
                                        If rsTot.EOF Then
                                            ' Creo el cubo
                                            StrSql = "INSERT INTO anrcubo(anrcabnro,anrcubmanual" & _
                                                ",anrrangnro,estrnro,facnro,tenro,Ternro,tipnro" & _
                                                ",anrcubvalor1,anrcubvalor2) VALUES (" & _
                                                rsAnrCab!anrcabnro & ",0," & rsRango!anrrangnro & "," & _
                                                rsEstructura!estrnro & "," & FactorTotalizador & "," & _
                                                rsEstructura!tenro & "," & rsAcumDiario!Ternro & ",1," & _
                                                cubvalor1 & "," & cubvalor2 & ")"
                                        Else
                                            StrSql = "UPDATE anrcubo SET" & _
                                                " anrcubvalor1 = " & Round(rsTot!anrcubvalor1 + cubvalor1, 2) & _
                                                " ,anrcubvalor2 = " & rs!anrcubvalor2 + cubvalor2 & _
                                                " WHERE anrcabnro = " & rsAnrCab!anrcabnro & _
                                                " AND facnro = " & FactorTotalizador & _
                                                " AND tenro = " & rsEstructura!tenro & _
                                                " AND estrnro = " & rsEstructura!estrnro & _
                                                " AND ternro = " & rsAcumDiario!Ternro & _
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
                                            " WHERE his_estructura.ternro = " & rsAcumDiario!Ternro & _
                                            " AND his_estructura.tenro = " & rsEstructura!tenro & _
                                            " AND his_estructura.htetdesde <= " & ConvFecha(rsRango!anrrangfechasta) & _
                                            " AND (his_estructura.htethasta >= " & ConvFecha(rsRango!anrrangfecdesde) & _
                                            " OR his_estructura.htethasta IS NULL) "
                                        OpenRecordset StrSql, objRs
                                        objRs.MoveLast
                                        
                                        If Not objRs.EOF Then
                                        
                                            StrSql = "SELECT * FROM anrcubo" & _
                                                " WHERE anrcabnro = " & rsAnrCab!anrcabnro & _
                                                " AND facnro = " & rsFactor!facnro & _
                                                " AND tenro = " & objRs!tenro & _
                                                " AND estrnro = " & objRs!estrnro & _
                                                " AND ternro = " & rsAcumDiario!Ternro & _
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
                                                    objRs!tenro & "," & rsAcumDiario!Ternro & ",1" & _
                                                    "," & cubvalor1 & "," & cubvalor2 & ")"
                                            Else
                                                StrSql = "UPDATE anrcubo SET" & _
                                                    " anrcubvalor1 = " & rs!anrcubvalor1 + cubvalor1 & _
                                                    " ,anrcubvalor2 = " & rs!anrcubvalor2 + cubvalor2 & _
                                                    " WHERE anrcabnro = " & rsAnrCab!anrcabnro & _
                                                    " AND facnro = " & rsFactor!facnro & _
                                                    " AND tenro = " & objRs!tenro & _
                                                    " AND estrnro = " & objRs!estrnro & _
                                                    " AND ternro = " & rsAcumDiario!Ternro & _
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
                                                    " AND ternro = " & rsAcumDiario!Ternro & _
                                                    " AND anrrangnro = " & rsRango!anrrangnro
                                                OpenRecordset StrSql, rsTot
                                                
                                                If rsTot.EOF Then
                                                    ' Creo el cubo
                                                    StrSql = "INSERT INTO anrcubo(anrcabnro,anrcubmanual" & _
                                                        ",anrrangnro,estrnro,facnro,tenro,Ternro,tipnro" & _
                                                        ",anrcubvalor1,anrcubvalor2) VALUES (" & _
                                                        rsAnrCab!anrcabnro & ",0," & rsRango!anrrangnro & "," & _
                                                        rsEstructura!estrnro & "," & FactorTotalizador & "," & _
                                                        rsEstructura!tenro & "," & rsAcumDiario!Ternro & ",1," & _
                                                        cubvalor1 & "," & cubvalor2 & ")"
                                                Else
                                                    StrSql = "UPDATE anrcubo SET" & _
                                                        " anrcubvalor1 = " & Round(rsTot!anrcubvalor1 + cubvalor1, 2) & _
                                                        " ,anrcubvalor2 = " & rsTot!anrcubvalor2 + cubvalor2 & _
                                                        " WHERE anrcabnro = " & rsAnrCab!anrcabnro & _
                                                        " AND facnro = " & FactorTotalizador & _
                                                        " AND tenro = " & rsEstructura!tenro & _
                                                        " AND estrnro = " & rsEstructura!estrnro & _
                                                        " AND ternro = " & rsAcumDiario!Ternro & _
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
                            
siguienteEstructura:
                            rsEstructura.MoveNext
                        Loop
                
                        rsRango.MoveNext
                    Loop
        
            ' Actualizo el progreso
            Progreso = Progreso + IncPorc
            StrSql = "UPDATE batch_proceso SET bprcprogreso = " & CInt(Progreso) & " WHERE bpronro = " & NroProceso
            objConn.Execute StrSql, , adExecuteNoRecords
        
            rsAcumDiario.MoveNext
        Loop
    
siguienteLegajo:
        rsFiltro.MoveNext
    Loop
    
    rsFactor.MoveNext
Loop
           
End Sub



Private Sub AcumuladoParcial(nro_analisis As Long, Filtrar As Boolean)

'Variables locales
Dim desde As Date
Dim hasta As Date
Dim horas As Single
Dim NroCab As Long
Dim Tercero As Long

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
Dim rsAcumParcial As New ADODB.Recordset
Dim rsAnrCab As New ADODB.Recordset
Dim rsFiltro As New ADODB.Recordset

'Código -------------------------------------------------------------------

'Abro la conexion para Nexus
OpenConnection strConexionNexus, objConnNexus

'Obtengo la cabecera
Call ObtenerCabecerayFiltro(nro_analisis, rsAnrCab, Filtrar, rsFiltro, cant_flt, Ok)

If Not Ok Then
    Exit Sub
End If

'Estoy reprocesando
Call PurgarCubo(nro_analisis, 5)


' Obtengo los rangos del analisis
StrSql = "SELECT * FROM anrrangofec" & _
         " WHERE anrrangofec.anrcabnro = " & rsAnrCab!anrcabnro
                        
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
    '/* Recorre para el analisis los acumulados Parcial de tipos de horas configurados */
    StrSql = "SELECT * FROM anrcab_fact" & _
        " INNER JOIN anrfact_ori ON anrfact_ori.facnro = anrcab_fact.facnro" & _
        " AND anrfact_ori.tipfacnro = 5" & _
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
    Call ObtenerLegajos(1, Filtrar, rsAnrCab!anrcabnro, rsFiltro, rsAnrCab!anrcabfecdesde, rsAnrCab!anrcabfechasta)
        
    'Recorro los acumulados Parciales que entran el el rango actual analizado
    Do While Not rsFactor.EOF
        
        'Busco si es factor totalizador
        StrSql = "SELECT * " & _
                " FROM   anrfact_tot, anrcab_fact" & _
                " WHERE  anrfact_tot.facnro = " & rsFactor!facnro & _
                " AND    anrcab_fact.facnro   = anrfact_tot.facnro " & _
                " AND    anrcab_fact.anrcabnro = " & rsAnrCab!anrcabnro
        OpenRecordset StrSql, rsFactorTotalizador
            
            
        'Busco si es factor totalizador
        'StrSql = "SELECT * FROM anrfact_tot" & _
        '        " WHERE facnro = " & rsFactor!facnro
        'OpenRecordset StrSql, rsFactorTotalizador
            
        If Not rsFactorTotalizador.EOF Then
            Totaliza = True
            ' codigo de factor con el cual se inserta en el cubo
            FactorTotalizador = rsFactorTotalizador!facnrotot
        Else
            Totaliza = False
            FactorTotalizador = 0
        End If
        
        Do While Not rsFiltro.EOF
            If (cant_flt > 0) And (rsFiltro!cant_te < cant_flt) Then
                GoTo siguienteLegajo
            End If
            
            '" INNER JOIN gti_achparc_estr ON gti_achparcial.achpnro = gti_achparc_estr.achpnro " & _

            StrSql = " SELECT * FROM gti_procacum" & _
                " INNER JOIN gti_cab ON gti_cab.gpanro = gti_procacum.gpanro " & _
                " INNER JOIN gti_det ON gti_det.cgtinro = gti_cab.cgtinro " & _
                " WHERE gti_procacum.gpadesde >= " & ConvFecha(Dia_Inicio_Per_Analizado) & _
                " AND gti_procacum.gpahasta <= " & ConvFecha(Dia_Fin_Per_Analizado) & _
                " AND gti_det.thnro = " & rsFactor!faccodorig & _
                " AND gti_cab.ternro = " & rsFiltro!Ternro & _
                " ORDER BY gti_cab.ternro"
            OpenRecordset StrSql, rsAcumParcial
            
            If Not rsAcumParcial.EOF Then
                IncPorc = CantRangos * (((100 / CantFactor) * (100 / rsAcumParcial.RecordCount)) / 100) / 100
            End If
        
            Do While Not rsAcumParcial.EOF
    
                            ' Obtengo las estructuras
                            Call ObtenerEstructuras(Filtrar, rsAcumParcial!Ternro, Dia_Inicio_Per_Analizado, Dia_Fin_Per_Analizado, rsEstructura)
    
                            'StrSql = "SELECT * FROM his_estructura" & _
                            '    " WHERE his_estructura.ternro = " & rsAcumParcial!Ternro & _
                            '    " AND his_estructura.htetdesde <= " & ConvFecha(Dia_Fin_Per_Analizado) & _
                            '    " AND (his_estructura.htethasta >= " & ConvFecha(Dia_Inicio_Per_Analizado) & _
                            '    " OR his_estructura.htethasta IS NULL)" & _
                            '    " ORDER BY ternro,tenro,estrnro"
                            'OpenRecordset StrSql, rsEstructura
        
       
                            If Not rsEstructura.EOF Then
                                TipoEstr = rsEstructura!tenro
                                EstrAct = rsEstructura!estrnro
                            End If
                        
                            Do While Not rsEstructura.EOF
        
                                If PrimerFactOri Then
                                    cantdiasper = DateDiff("d", rsRango!anrrangfecdesde, rsRango!anrrangfechasta) + 1
                                    monto_total = 0
                                    cant_total = 0
                                    cant_saldo = 0
                                    PrimerFactOri = False
                                End If
                                               
                                '/* Acumulo por Factor */
                                'monto_total = monto_total + rsAcumParcial!achpcanthoras
                                monto_total = monto_total + rsAcumParcial!dgticant
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
                                    
                                    If Last_OF_Factor() Or Last_OF_estrnro() Then
                                        
                                        cubvalor1 = monto_total * porcentaje / 100
                                        cubvalor2 = cant_total * porcentaje / 100
                                        
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
                                                rsEstructura!tenro & "," & rsAcumParcial!Ternro & ",1"
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
                                                " AND ternro = " & rsAcumParcial!Ternro & _
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
                                                " AND ternro = " & rsAcumParcial!Ternro & _
                                                " AND anrrangnro = " & rsRango!anrrangnro
                                            OpenRecordset StrSql, rsTot
                                            
                                            If rsTot.EOF Then
                                                ' Creo el cubo
                                                StrSql = "INSERT INTO anrcubo(anrcabnro,anrcubmanual" & _
                                                    ",anrrangnro,estrnro,facnro,tenro,Ternro,tipnro" & _
                                                    ",anrcubvalor1,anrcubvalor2) VALUES (" & _
                                                    rsAnrCab!anrcabnro & ",0," & rsRango!anrrangnro & "," & _
                                                    rsEstructura!estrnro & "," & FactorTotalizador & "," & _
                                                    rsEstructura!tenro & "," & rsAcumParcial!Ternro & ",1," & _
                                                    cubvalor1 & "," & cubvalor2 & ")"
                                            Else
                                                StrSql = "UPDATE anrcubo SET" & _
                                                    " anrcubvalor1 = " & Round(rsTot!anrcubvalor1 + cubvalor1, 2) & _
                                                    " ,anrcubvalor2 = " & rs!anrcubvalor2 + cubvalor2 & _
                                                    " WHERE anrcabnro = " & rsAnrCab!anrcabnro & _
                                                    " AND facnro = " & FactorTotalizador & _
                                                    " AND tenro = " & rsEstructura!tenro & _
                                                    " AND estrnro = " & rsEstructura!estrnro & _
                                                    " AND ternro = " & rsAcumParcial!Ternro & _
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
                                                " WHERE his_estructura.ternro = " & rsAcumParcial!Ternro & _
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
                                                    " AND ternro = " & rsAcumParcial!Ternro & _
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
                                                        objRs!tenro & "," & rsAcumParcial!Ternro & ",1" & _
                                                        "," & cubvalor1 & "," & cubvalor2 & ")"
                                                Else
                                                    StrSql = "UPDATE anrcubo SET" & _
                                                        " anrcubvalor1 = " & rs!anrcubvalor1 + cubvalor1 & _
                                                        " ,anrcubvalor2 = " & rs!anrcubvalor2 + cubvalor2 & _
                                                        " WHERE anrcabnro = " & rsAnrCab!anrcabnro & _
                                                        " AND facnro = " & rsFactor!facnro & _
                                                        " AND tenro = " & objRs!tenro & _
                                                        " AND estrnro = " & objRs!estrnro & _
                                                        " AND ternro = " & rsAcumParcial!Ternro & _
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
                                                        " AND ternro = " & rsAcumParcial!Ternro & _
                                                        " AND anrrangnro = " & rsRango!anrrangnro
                                                    OpenRecordset StrSql, rsTot
                                                    
                                                    If rsTot.EOF Then
                                                        ' Creo el cubo
                                                        StrSql = "INSERT INTO anrcubo(anrcabnro,anrcubmanual" & _
                                                            ",anrrangnro,estrnro,facnro,tenro,Ternro,tipnro" & _
                                                            ",anrcubvalor1,anrcubvalor2) VALUES (" & _
                                                            rsAnrCab!anrcabnro & ",0," & rsRango!anrrangnro & "," & _
                                                            rsEstructura!estrnro & "," & FactorTotalizador & "," & _
                                                            rsEstructura!tenro & "," & rsAcumParcial!Ternro & ",1," & _
                                                            cubvalor1 & "," & cubvalor2 & ")"
                                                    Else
                                                        StrSql = "UPDATE anrcubo SET" & _
                                                            " anrcubvalor1 = " & Round(rsTot!anrcubvalor1 + cubvalor1, 2) & _
                                                            " ,anrcubvalor2 = " & rs!anrcubvalor2 + cubvalor2 & _
                                                            " WHERE anrcabnro = " & rsAnrCab!anrcabnro & _
                                                            " AND facnro = " & FactorTotalizador & _
                                                            " AND tenro = " & rsEstructura!tenro & _
                                                            " AND estrnro = " & rsEstructura!estrnro & _
                                                            " AND ternro = " & rsAcumParcial!Ternro & _
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
                    
                rsAcumParcial.MoveNext
            Loop

siguienteLegajo:
        
            rsFiltro.MoveNext
        Loop
        
        ' voy nuevamente al primer legajo del filtro
        rsFiltro.MoveFirst
        
        rsFactor.MoveNext
    Loop
           
           
    ' Actualizo el progreso
    Progreso = Progreso + IncPorc
    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & CInt(Progreso) & " WHERE bpronro = " & NroProceso
    objConn.Execute StrSql, , adExecuteNoRecords
           
    rsRango.MoveNext
Loop
End Sub


Private Sub Licencias(nro_analisis As Long, Filtrar As Boolean)

'Variables locales
Dim desde As Date
Dim hasta As Date
Dim horas As Single
Dim NroCab As Long
Dim Tercero As Long

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
OpenConnection strConexionNexus, objConnNexus

'Obtengo la cabecera
Call ObtenerCabecerayFiltro(nro_analisis, rsAnrCab, Filtrar, rsFiltro, cant_flt, Ok)

If Not Ok Then
    Exit Sub
End If

'Estoy reprocesando
Call PurgarCubo(nro_analisis, 6)


' Obtengo los rangos del analisis
StrSql = "SELECT * FROM anrrangofec" & _
         " WHERE anrrangofec.anrcabnro = " & rsAnrCab!anrcabnro
                        
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
    Call ObtenerLegajos(1, Filtrar, rsAnrCab!anrcabnro, rsFiltro, rsAnrCab!anrcabfecdesde, rsAnrCab!anrcabfechasta)
        
    'Recorro las Licencias que entran el el rango actual analizado
    Do While Not rsFactor.EOF
        
        'Busco si es factor totalizador
        StrSql = "SELECT * " & _
                " FROM   anrfact_tot, anrcab_fact" & _
                " WHERE  anrfact_tot.facnro = " & rsFactor!facnro & _
                " AND    anrcab_fact.facnro   = anrfact_tot.facnro " & _
                " AND    anrcab_fact.anrcabnro = " & rsAnrCab!anrcabnro
        OpenRecordset StrSql, rsFactorTotalizador
            
            
        'Busco si es factor totalizador
        'StrSql = "SELECT * FROM anrfact_tot" & _
        '        " WHERE facnro = " & rsFactor!facnro
        'OpenRecordset StrSql, rsFactorTotalizador
            
        If Not rsFactorTotalizador.EOF Then
            Totaliza = True
            ' codigo de factor con el cual se inserta en el cubo
            FactorTotalizador = rsFactorTotalizador!facnrotot
        Else
            Totaliza = False
            FactorTotalizador = 0
        End If
            
        ' voy al primer legajo nuevamente
        rsFiltro.MoveFirst
        
        Do While Not rsFiltro.EOF
            If (cant_flt > 0) And (rsFiltro!cant_te < cant_flt) Then
                GoTo siguienteLegajo
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
                                    
                                    If Last_OF_Factor() Or Last_OF_estrnro() Then
                                        
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
                                        
                                        'FZG 25/07/2003
                                        'Actualizo Totalizador
                                        If Totaliza Then
                                            StrSql = "SELECT * FROM anrcubo" & _
                                                " WHERE anrcabnro = " & rsAnrCab!anrcabnro & _
                                                " AND facnro = " & FactorTotalizador & _
                                                " AND tenro = " & rsEstructura!tenro & _
                                                " AND estrnro = " & rsEstructura!estrnro & _
                                                " AND ternro = " & rsLicencias!Ternro & _
                                                " AND anrrangnro = " & rsRango!anrrangnro
                                            OpenRecordset StrSql, rsTot
                                            
                                            If rsTot.EOF Then
                                                ' Creo el cubo
                                                StrSql = "INSERT INTO anrcubo(anrcabnro,anrcubmanual" & _
                                                    ",anrrangnro,estrnro,facnro,tenro,Ternro,tipnro" & _
                                                    ",anrcubvalor1,anrcubvalor2) VALUES (" & _
                                                    rsAnrCab!anrcabnro & ",0," & rsRango!anrrangnro & "," & _
                                                    rsEstructura!estrnro & "," & FactorTotalizador & "," & _
                                                    rsEstructura!tenro & "," & rsLicencias!Ternro & ",1," & _
                                                    cubvalor1 & "," & cubvalor2 & ")"
                                            Else
                                                StrSql = "UPDATE anrcubo SET" & _
                                                    " anrcubvalor1 = " & Round(rsTot!anrcubvalor1 + cubvalor1, 2) & _
                                                    " ,anrcubvalor2 = " & rs!anrcubvalor2 + cubvalor2 & _
                                                    " WHERE anrcabnro = " & rsAnrCab!anrcabnro & _
                                                    " AND facnro = " & FactorTotalizador & _
                                                    " AND tenro = " & rsEstructura!tenro & _
                                                    " AND estrnro = " & rsEstructura!estrnro & _
                                                    " AND ternro = " & rsLicencias!Ternro & _
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
                                                
                                                'FZG 25/07/2003
                                                'Actualizo Totalizador
                                                If Totaliza Then
                                                    StrSql = "SELECT * FROM anrcubo" & _
                                                        " WHERE anrcabnro = " & rsAnrCab!anrcabnro & _
                                                        " AND facnro = " & FactorTotalizador & _
                                                        " AND tenro = " & rsEstructura!tenro & _
                                                        " AND estrnro = " & rsEstructura!estrnro & _
                                                        " AND ternro = " & rsLicencias!Ternro & _
                                                        " AND anrrangnro = " & rsRango!anrrangnro
                                                    OpenRecordset StrSql, rsTot
                                                    
                                                    If rsTot.EOF Then
                                                        ' Creo el cubo
                                                        StrSql = "INSERT INTO anrcubo(anrcabnro,anrcubmanual" & _
                                                            ",anrrangnro,estrnro,facnro,tenro,Ternro,tipnro" & _
                                                            ",anrcubvalor1,anrcubvalor2) VALUES (" & _
                                                            rsAnrCab!anrcabnro & ",0," & rsRango!anrrangnro & "," & _
                                                            rsEstructura!estrnro & "," & FactorTotalizador & "," & _
                                                            rsEstructura!tenro & "," & rsLicencias!Ternro & ",1," & _
                                                            cubvalor1 & "," & cubvalor2 & ")"
                                                    Else
                                                        StrSql = "UPDATE anrcubo SET" & _
                                                            " anrcubvalor1 = " & Round(rsTot!anrcubvalor1 + cubvalor1, 2) & _
                                                            " ,anrcubvalor2 = " & rs!anrcubvalor2 + cubvalor2 & _
                                                            " WHERE anrcabnro = " & rsAnrCab!anrcabnro & _
                                                            " AND facnro = " & FactorTotalizador & _
                                                            " AND tenro = " & rsEstructura!tenro & _
                                                            " AND estrnro = " & rsEstructura!estrnro & _
                                                            " AND ternro = " & rsLicencias!Ternro & _
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
                    
                rsLicencias.MoveNext
            Loop
                    
siguienteLegajo:
            rsFiltro.MoveNext
        Loop
        
        rsFactor.MoveNext
    Loop
           
           
    ' Actualizo el progreso
    Progreso = Progreso + IncPorc
    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & CInt(Progreso) & " WHERE bpronro = " & NroProceso
    objConn.Execute StrSql, , adExecuteNoRecords
           
    rsRango.MoveNext
Loop
End Sub

