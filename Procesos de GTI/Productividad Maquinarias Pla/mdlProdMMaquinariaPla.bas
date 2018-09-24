Attribute VB_Name = "mdlProdMMaquinariaPla"
Option Explicit

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Const Version = "1.00" 'Version inicial
'Const FechaVersion = "16/07/2014"
'CAS-24475 - PLA - Nuevo Reporte de Produccion Pulverizadoras - LED


Const Version = "1.01"
Const FechaVersion = "29/12/2014"
'CAS-24475 - PLA - Nuevo Reporte de Produccion Pulverizadoras [Entrega 3] - LED - cambio en el porcentaje de contribucion

'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Dim fs
Dim Flog
Dim FDesde As Date
Dim FHasta As Date
Global Fecha_Inicio As Date
Global CEmpleadosAProc As Long
Global IncPorc As Double
Global Progreso As Double
Global TiempoInicialProceso As Long
Global totalEmpleados As Long
Global TiempoAcumulado As Long
Global cantRegistros As Long
Dim l_iduser
Dim l_estrnro1
Dim l_estrnro2
Dim l_estrnro3




Sub Main()

Dim Archivo As String
Dim pos As Integer
Dim strcmdLine  As String
Dim rs As New ADODB.Recordset

Dim NroProceso As Long
Dim NroReporte As Long
Dim StrParametros As String

Dim rs_Batch_Proceso As New ADODB.Recordset
Dim rs_His_Batch_Proceso As New ADODB.Recordset

Dim PID As String
Dim ArrParametros


    strcmdLine = Command()
    ArrParametros = Split(strcmdLine, " ", -1)
    If UBound(ArrParametros) > 1 Then
        If IsNumeric(ArrParametros(0)) Then
            NroProceso = ArrParametros(0)
            Etiqueta = ArrParametros(1)
            EncriptStrconexion = CBool(ArrParametros(2))
            c_seed = ArrParametros(2)
        Else
            Exit Sub
        End If
    Else
        If UBound(ArrParametros) > 0 Then
            If IsNumeric(ArrParametros(0)) Then
                NroProceso = ArrParametros(0)
                Etiqueta = ArrParametros(1)
            Else
                Exit Sub
            End If
        Else
            If IsNumeric(strcmdLine) Then
                NroProceso = strcmdLine
            Else
                Exit Sub
            End If
        End If
    End If
    
    
    ' carga las configuraciones basicas, formato de fecha, string de conexion,
    ' tipo de BD y ubicacion del archivo de log
    Call CargarConfiguracionesBasicas

    'Creo el archivo de texto del desglose
    Archivo = PathFLog & "ProdMaqPla_" & CStr(NroProceso) & ".log"

    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Archivo, True)


    'Abro la conexion
    On Error Resume Next
    OpenConnection strconexion, objConn
    If Err.Number <> 0 Or Error_Encrypt Then
        Flog.writeline Espacios(Tabulador * 0) & "Problemas en la conexion"
        Exit Sub
    End If
    On Error Resume Next
    OpenConnection strconexion, objConnProgreso
    If Err.Number <> 0 Or Error_Encrypt Then
        Flog.writeline Espacios(Tabulador * 0) & "Problemas en la conexion"
        Exit Sub
    End If

    On Error GoTo ce


    'Obtengo el Process ID
    PID = GetCurrentProcessId
    Flog.writeline "-------------------------------------------------"
    Flog.writeline "Version                  : " & Version
    Flog.writeline "Fecha Ultima Modificacion: " & FechaVersion
    Flog.writeline "PID                      : " & PID
    Flog.writeline "-------------------------------------------------"
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 0) & "Inicio :" & Now
    
    'FGZ - 05/08/2009 --------- Control de versiones ------
    Version_Valida = ValidarV(Version, 1, TipoBD)
    If Not Version_Valida Then
        'Actualizo el progreso
        StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcprogreso = 0, bprcestado = 'Error de Version', bprcpid = " & PID & " WHERE bpronro = " & NroProceso
        objConnProgreso.Execute StrSql, , adExecuteNoRecords
        Flog.writeline
        GoTo Final
    End If
    'FGZ - 05/08/2009 --------- Control de versiones ------

    'Cambio el estado del proceso a Procesando
    StrSql = "UPDATE batch_proceso SET bprcpid = " & PID & ", bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcestado = 'Procesando', bprcprogreso = 0 WHERE bpronro = " & NroProceso
    objConn.Execute StrSql, , adExecuteNoRecords

    Flog.writeline "Levanta Proceso y Setea Parámetros:  " & " " & Now

    'levanto los parametros del proceso
    StrParametros = ""

    StrSql = "SELECT bprcfecdesde,bprcfechasta,bprcparam,iduser FROM batch_proceso WHERE bpronro = " & NroProceso
    OpenRecordset StrSql, rs
    If Not rs.EOF Then
        Flog.writeline "Inicio de Reporte Produccion de maquinaria: " & Now
        Call Rep_produccion_Maquinaria(rs!bprcparam, NroProceso)
    Else
        Exit Sub
    End If
        
    ' poner el bprcestado en procesado
    StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado', bprcprogreso = 100 WHERE bpronro = " & NroProceso
    objConn.Execute StrSql, , adExecuteNoRecords

    If objConn.State = adStateOpen Then objConn.Close
    If objConnProgreso.State = adStateOpen Then objConnProgreso.Close
    
    Set rs_Batch_Proceso = Nothing
    Set rs_His_Batch_Proceso = Nothing

Final:
    Flog.writeline Espacios(Tabulador * 0) & "Fin de Reporte de produccion de maquinarias: " & " " & Now
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 0) & "---------------------------------------------------------------------------------"
    Flog.writeline Espacios(Tabulador * 0) & "---------------------------------------------------------------------------------"
    Flog.Close
Exit Sub

ce:
    Flog.writeline Espacios(Tabulador * 0) & "Reporte abortado por Error:" & " " & Now
    Flog.writeline Espacios(Tabulador * 0) & "Reporte abortado por :" & Err.Description
    Flog.writeline Espacios(Tabulador * 0) & "Ultimo SQL " & StrSql
    StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Error' WHERE bpronro = " & NroProceso
    objConn.Execute StrSql, , adExecuteNoRecords
End Sub
Private Sub Rep_produccion_Maquinaria(parametros As String, NroProceso As Long)

Dim rs As New ADODB.Recordset
Dim rsEmp As New ADODB.Recordset
Dim rsEstruc As New ADODB.Recordset
Dim historicoDesc As String
Dim hsdiasHab As Double
Dim ArrParametros
Dim ternro As String
Dim fecdesde As String
Dim fechasta As String
Dim estrnroEmpresa As String
Dim tipoMaquinaria As String
Dim emplogo As String
Dim mes As Integer
Dim anio As String
Dim listaEmpleados As String
Dim prom As Double
Dim prom1 As Double
Dim prom2 As Double
Dim i, j As Integer
Dim promediar As Long
Dim tipomaq As String
Dim estrnro1 As String
Dim estrnro2 As String
Dim ArrDatos(27, 16)  'se guardan los datos del reporte con la forma ArrDatos(fila, columna)
Dim thnormal As String
Dim thPagExtras As String
Dim thPagSinExtras As String
Dim thHsFeriado As String
Dim thLicPagas As String
Dim thLicEnfermedad As String
Dim thAccidente As String
Dim thHsImproductivas As String
Dim th50 As String
Dim th100 As String
Dim fecDesdeBucle As String
Dim fecHastaBucle As String
Dim empresa As String
Dim cgrnro As String
Dim logoAncho As String
Dim logoAlto As String
Dim fila As Integer
Dim tipimnro As String

'======================================================================================================
'SE VALIDA Y LEVANTAN PARAMETROS.
'======================================================================================================
        
    Flog.writeline ""
    Flog.writeline "Parametros:" & parametros
    Flog.writeline ""
    
    ' parametros(0) --> legdesde
    ' parametros(1) --> leghasta
    ' parametros(2) --> estado
    ' parametros(3) --> tenro1
    ' parametros(4) --> estrnro1
    ' parametros(5) --> tenro2
    ' parametros(6) --> estrnro2
    ' parametros(7) --> tenro3
    ' parametros(8) --> estrnro3
    ' parametros(9) --> anio
    ' parametros(10) --> empresa
    ' parametros(11) --> 1 pulverizadoras, 2 sembradoras
    
    ArrParametros = Split(parametros, "@")
    estrnro1 = ArrParametros(4)
    estrnro2 = ArrParametros(6)
    anio = Year(ArrParametros(9))
    fecdesde = ArrParametros(9)
    fechasta = ArrParametros(10)
    estrnroEmpresa = ArrParametros(11)
    tipomaq = ArrParametros(12)
    promediar = CLng(Month(fechasta)) - CLng(Month(fecdesde)) + 1
    
    Flog.writeline "Busco la estructura configurada en el confrep"
    StrSql = " SELECT confnrocol, confval, estrdabr FROM confrep " & _
             " INNER JOIN estructura ON estructura.estrnro = confrep.confval2 " & _
             " WHERE repnro = 443 "
    If CLng(tipomaq) = 1 Then
        StrSql = StrSql & " AND confnrocol in (1,2) "
    End If
        
    If CLng(tipomaq) = 2 Then
        StrSql = StrSql & " AND confnrocol in (3,4) "
    End If
    
    OpenRecordset StrSql, rs
    Do While Not rs.EOF
        Select Case CLng(rs!confnrocol)
            Case 1 'sector sin uso por ahora
            
            Case 2 'unidad de negocio
                tipoMaquinaria = rs!estrdabr
                Flog.writeline "Estructura encontrada: " & tipoMaquinaria
            
            Case 3 'sector sin uso por ahora
            
            Case 4 'unidad de negocio
                tipoMaquinaria = rs!estrdabr
                Flog.writeline "Estructura encontrada: " & tipoMaquinaria
                
        End Select
        rs.MoveNext
    Loop
                
    thnormal = "0"
    thPagExtras = "0"
    thPagSinExtras = "0"
    thHsFeriado = "0"
    thLicPagas = "0"
    thLicEnfermedad = "0"
    thAccidente = "0"
    thHsImproductivas = "0"
    th50 = "0"
    th100 = "0"
    StrSql = " SELECT confnrocol, confval, confval2 FROM confrep WHERE repnro = 443 "
    OpenRecordset StrSql, rs
    Do While Not rs.EOF
        Select Case CLng(rs!confnrocol)
            Case 5 'valor promedio de hora diaria
                hsdiasHab = rs!confval2
                Flog.writeline "Cantidad de horas habiles encontrada: " & hsdiasHab
            
            Case 6 'lista de horas para la fila 4 - Total de Personal por cantidad de Hs Reales
                thnormal = thnormal & "," & rs!confval2
                Flog.writeline "Lista de horas normales encontrada: " & thnormal
        
            Case 7 'lista de horas para la fila 6 - Total Horas Pagadas c/ Extras
                thPagExtras = thPagExtras & "," & rs!confval2
                Flog.writeline "Lista de Horas Pagadas c/ Extras encontrada: " & thPagExtras
        
            Case 8 'lista de horas para la fila 7 - Total de Horas Pagadas s/ Extras
                thPagSinExtras = thPagSinExtras & "," & rs!confval2
                Flog.writeline "Lista de Total de Horas Pagadas s/ Extras encontrada: " & thPagExtras
        
            Case 9 'lista de horas para la fila 10 - Feriados Pagos
                thHsFeriado = thHsFeriado & "," & rs!confval2
                Flog.writeline "Lista de Feriados Pagos: " & thHsFeriado
            
            Case 10 'lista de horas para la fila 11 - Licencias Pagas
                thLicPagas = thLicPagas & "," & rs!confval2
                Flog.writeline "Lista Hs Licencias Pagas: " & thLicPagas

            Case 11 'lista de horas para la fila 12 - Enfermedad
                thLicPagas = thLicPagas & "," & rs!confval2
                Flog.writeline "Lista de hs Enfermedad: " & thLicEnfermedad
            
            Case 12 'lista de horas para la fila 13 - Accidente
                thAccidente = thAccidente & "," & rs!confval2
                Flog.writeline "Lista de hs Accidente: " & thAccidente

            Case 13 'lista de horas para la fila 14 - Horas Improductivas
                thHsImproductivas = thHsImproductivas & "," & rs!confval2
                Flog.writeline "Lista de hs Improductivas: " & thHsImproductivas

            Case 14 'lista de horas para la fila 18 - Codigo de escala Categoria
                cgrnro = rs!confval
                Flog.writeline "Codigo de escala Categoria obtenido: " & cgrnro

            Case 15 'lista de horas para la fila 24 - Horas al 50%
                th50 = th50 & "," & rs!confval2
                Flog.writeline "Lista de hs al 50: " & th50
            
            Case 16 'lista de horas para la fila 25 - Horas al 100%
                th100 = th100 & "," & rs!confval2
                Flog.writeline "Lista de hs al 100: " & th100

            Case 17 'tipo de imagen para el logo
                tipimnro = rs!confval
                Flog.writeline "Tipo de imagen configurado: " & tipimnro

        End Select
        rs.MoveNext
    Loop
    
    historicoDesc = NroProceso & " - Rep Produccion " & tipoMaquinaria & ". Año: " & anio
    
    Call obtenerEmpleados(NroProceso, listaEmpleados)
    
    If listaEmpleados <> "0" Then
    
        '---------------------------------------------------------------------------------------------------------------------
        'borro el hsitorico correspondiente al año que se quiere ejecutar
        StrSql = " SELECT bpronro FROM gti_rep_maquinarias WHERE anio = " & anio & " AND tipomaq = " & tipomaq
        OpenRecordset StrSql, rs
        If Not rs.EOF Then
            'borro el detalle
            StrSql = " DELETE FROM gti_rep_maquinarias_det WHERE bpronro = " & rs!bpronro
            objConn.Execute StrSql, , adExecuteNoRecords
            
            'borro la cabecera
            StrSql = " DELETE FROM gti_rep_maquinarias WHERE anio = " & anio & " AND tipomaq = " & tipomaq
            objConn.Execute StrSql, , adExecuteNoRecords
            
            Flog.writeline "Borrado el historico para el año: " & anio & "."
        End If
        
        
        '---------------------------------------------------------------------------------------------------------------------
        'busco el logo de la empresa
        StrSql = " SELECT ter_imag.terimnombre, tipoimag.tipimdire, tipoimag.tipimanchodef, tipoimag.tipimaltodef, empresa.empnom " & _
                 " From empresa " & _
                 " LEFT JOIN ter_imag ON ter_imag.ternro = empresa.ternro AND ter_imag.tipimnro = " & tipimnro & _
                 " LEFT JOIN tipoimag ON  tipoimag.tipimnro = ter_imag.tipimnro AND tipoimag.tipimnro = " & tipimnro & _
                 " WHERE empresa.estrnro = " & estrnroEmpresa
        OpenRecordset StrSql, rs
        If Not rs.EOF Then
            empresa = rs!terimnombre
            emplogo = rs!tipimdire & rs!terimnombre
            logoAncho = rs!tipimanchodef
            logoAlto = rs!tipimaltodef
            Flog.writeline "Logo de la Empresa encontrado."
        Else
            logoAncho = 0
            logoAlto = 0
            emplogo = ""
            Flog.writeline "Logo de la Empresa no encontrado."
        End If
        
        '----------------------------------------------------------------------------------------------------------------
        For i = 1 To 27
            For j = 1 To 16
                ArrDatos(i, j) = 0
            Next
        Next
        
        For i = 1 To 27
            
            ArrDatos(i, 1) = buscarHistorico(CLng(anio) - 2, i, tipomaq) 'promedio 2 años antes - MAQUINAS PLANIFICADAS
            ArrDatos(i, 2) = buscarHistorico(CLng(anio) - 1, i, tipomaq) 'promedio 1 año antes - MAQUINAS PLANIFICADAS
            'Fila 1 - MAQUINAS PLANIFICADAS
            'Fila 2 - MAQUINAS ENTREGADAS
            'Fila 3 - Total de Personal de Producción
            'Fila 4 - Total de Personal por cantidad de Hs Reales
            'Fila 5 - hs.Pdio.por Empleado
            'Fila 6 - Total de Horas Teoricas
            'Fila 7 - Total Horas Pagadas c/ Extras
            'Fila 8 - Total de Horas Pagadas s/ Extras
            'Fila 9 - Horas Desviadas
            'Fila 10 - Feriados Pagos
            'Fila 11 - Licencias Pagas
            'Fila 12 - Enfermedad
            'Fila 13 - Accidente
            'Fila 14 - Horas Improductivas
            'Fila 15 - Horas Netas de Produccion
            'Fila 16 - Horas por Máquina Producida
            'Fila 17 - Horas Standard de fabricacion
            'Fila 18 - Costo pdio.Gral Horas c/ Cs.Soc.
            'Fila 19 - Costo pdio. Hora 50 %
            'Fila 20 - Costo pdio. Hora 100 %
            'Fila 21 - Variación en Horas
            'Fila 22 - Variación en Miles de Pesos
            'Fila 23 - Variación en Porcentajes
            'Fila 24 - Horas al 50%
            'Fila 25 - Horas al 100%
            'Fila 26 - Total Horas Extras
            'Fila 27 - COSTO CON CARGAS SOCIALES
        Next

        '----------------------------------------------------------------------------------------------------------------
        Flog.writeline "Calculo de progreso"
        Progreso = 0
        If promediar > 0 Then
            IncPorc = (100 / promediar)
        Else
            IncPorc = 100
        End If
        TiempoInicialProceso = GetTickCount
        
        
        Flog.writeline "Inicializacion de Variables"
        'Este bucle calcula los datos hasta el mes que se tira el reporte porque tiene que promediar mes a mes
        Flog.writeline "Comienzo del analisis por mes"
        For mes = CLng(Month(fecdesde)) To CLng(Month(fechasta))
            Flog.writeline "Analizando mes: " & mes & " año: " & anio
            Progreso = Progreso + IncPorc
            fecDesdeBucle = "01/" & Right("00" & mes, 2) & "/" & anio
            fecHastaBucle = CStr(DateSerial(anio, mes + 1, 0))
            
            
            'busco los datos para todos los meses
            Flog.writeline "Buscando datos - MAQUINAS PLANIFICADAS"
            ArrDatos(1, mes + 2) = maquinasPlanificadas(mes, anio, tipomaq)
            ArrDatos(1, 15) = ArrDatos(1, 15) + ArrDatos(1, mes + 2)
            'ArrDatos(1, 16) = ArrDatos(1, 16) + ArrDatos(1, mes + 2)
            
            Flog.writeline "Buscando datos - MAQUINAS ENTREGADAS A VTAS"
            ArrDatos(2, mes + 2) = maquinasEntregadas(mes, anio, tipomaq)
            ArrDatos(2, 15) = ArrDatos(2, 15) + ArrDatos(2, mes + 2)
            'ArrDatos(2, 16) = ArrDatos(2, 16) + ArrDatos(2, mes + 2)
            
            Flog.writeline "Buscando datos - Total de Personal de Producción"
            ArrDatos(3, mes + 2) = empleadosEstructura(estrnro1, estrnro2, fecDesdeBucle, fecHastaBucle)
            ArrDatos(3, 16) = ArrDatos(3, 16) + ArrDatos(3, mes + 2)
            'totalFila3 = totalFila3 + ArrDatos(3, mes + 2)
            
            Flog.writeline "Buscando datos - Total de Personal por cantidad de Hs Reales"
            ArrDatos(4, mes + 2) = obtenerhoras(thnormal, listaEmpleados, fecDesdeBucle, fecHastaBucle, CDbl(hsdiasHab))
            ArrDatos(4, 16) = ArrDatos(4, 16) + ArrDatos(4, mes + 2)
            'totalFila4 = totalFila4 + ArrDatos(4, mes + 2)

            Flog.writeline "Buscando datos - Total de Horas Teoricas"
            ArrDatos(6, mes + 2) = ArrDatos(3, mes + 2) * CDbl(hsdiasHab) * diashabiles(fecDesdeBucle, fecHastaBucle)
            ArrDatos(6, 16) = ArrDatos(6, 16) + ArrDatos(6, mes + 2)
            'totalFila6 = totalFila6 + ArrDatos(6, mes + 2)
            
            Flog.writeline "Buscando datos - Total Horas Pagadas c/ Extras"
            ArrDatos(7, mes + 2) = obtenerhorasSinFormula(thPagExtras, listaEmpleados, fecDesdeBucle, fecHastaBucle)
            ArrDatos(7, 16) = ArrDatos(7, 16) + ArrDatos(7, mes + 2)
            
            Flog.writeline "Buscando datos - Total de Horas Pagadas s/ Extras"
            ArrDatos(8, mes + 2) = obtenerhorasSinFormula(thPagSinExtras, listaEmpleados, fecDesdeBucle, fecHastaBucle)
            ArrDatos(8, 16) = ArrDatos(8, 16) + ArrDatos(8, mes + 2)
            
            Flog.writeline "Buscando datos - Hs. Pdio. por Empleado"
            If ArrDatos(3, mes + 2) > 0 Then
                ArrDatos(5, mes + 2) = ArrDatos(8, mes + 2) / ArrDatos(3, mes + 2)
                ArrDatos(5, 16) = ArrDatos(5, 16) + ArrDatos(5, mes + 2)
                'totalFila5 = totalFila5 + ArrDatos(5, mes + 2)
            End If

            Flog.writeline "Buscando datos - Horas Desviadas"
            ArrDatos(9, mes + 2) = ArrDatos(7, mes + 2) - ArrDatos(6, mes + 2)
            ArrDatos(9, 16) = ArrDatos(9, 16) + ArrDatos(9, mes + 2)
            'totalFila9 = totalFila9 + ArrDatos(9, mes + 2)
                        
            Flog.writeline "Buscando datos - Feriados Pagos"
            ArrDatos(10, mes + 2) = obtenerhorasSinFormula(thHsFeriado, listaEmpleados, fecDesdeBucle, fecHastaBucle)
            ArrDatos(10, 16) = ArrDatos(10, 16) + ArrDatos(10, mes + 2)
            'totalFila10 = totalFila10 + ArrDatos(10, mes + 2)
            
            Flog.writeline "Buscando datos - Licencias Pagas"
            ArrDatos(11, mes + 2) = obtenerhorasSinFormula(thLicPagas, listaEmpleados, fecDesdeBucle, fecHastaBucle)
            ArrDatos(11, 16) = ArrDatos(11, 16) + ArrDatos(11, mes + 2)
            'totalFila11 = totalFila11 + ArrDatos(11, mes + 2)

            Flog.writeline "Buscando datos - Enfermedad"
            ArrDatos(12, mes + 2) = obtenerhorasSinFormula(thLicEnfermedad, listaEmpleados, fecDesdeBucle, fecHastaBucle)
            ArrDatos(12, 16) = ArrDatos(12, 16) + ArrDatos(12, mes + 2)
            'totalFila12 = totalFila12 + ArrDatos(12, mes + 2)

            Flog.writeline "Buscando datos - Accidente"
            ArrDatos(13, mes + 2) = obtenerhorasSinFormula(thAccidente, listaEmpleados, fecDesdeBucle, fecHastaBucle)
            ArrDatos(13, 16) = ArrDatos(13, 16) + ArrDatos(13, mes + 2)
            'totalFila13 = totalFila13 + ArrDatos(13, mes + 2)

            Flog.writeline "Buscando datos - Horas Improductivas"
            ArrDatos(14, mes + 2) = obtenerhorasSinFormula(thHsImproductivas, listaEmpleados, fecDesdeBucle, fecHastaBucle)
            ArrDatos(14, 16) = ArrDatos(14, 16) + ArrDatos(14, mes + 2)
            'totalFila14 = totalFila14 + ArrDatos(14, mes + 2)

            Flog.writeline "Buscando datos - Horas Netas de Producción"
            ArrDatos(15, mes + 2) = ArrDatos(7, mes + 2) - ArrDatos(10, mes + 2) - ArrDatos(11, mes + 2) - ArrDatos(12, mes + 2) - ArrDatos(13, mes + 2) - ArrDatos(14, mes + 2)
            ArrDatos(15, 16) = ArrDatos(15, 16) + ArrDatos(15, mes + 2)
            'totalFila15 = totalFila15 + ArrDatos(15, mes + 2)

            Flog.writeline "Buscando datos - Horas por Máquina Producida"
            If ArrDatos(2, mes + 2) > 0 Then
                ArrDatos(16, mes + 2) = ArrDatos(15, mes + 2) / ArrDatos(2, mes + 2)
                ArrDatos(16, 16) = ArrDatos(16, 16) + ArrDatos(16, mes + 2)
                'totalFila16 = totalFila16 + ArrDatos(16, mes + 2)
            End If
            

            Flog.writeline "Buscando datos - Standard de Fabricación"
            ArrDatos(17, mes + 2) = HsStandard(mes, anio, tipomaq)
            ArrDatos(17, 16) = ArrDatos(17, 16) + ArrDatos(17, mes + 2)
            'totalFila17 = totalFila17 + ArrDatos(17, mes + 2)

            Flog.writeline "Buscando datos - Costo pdio.Gral Horas c/ Cs.Soc."
            If ArrDatos(3, mes + 2) > 0 Then
                ArrDatos(18, mes + 2) = (escalaCategoria(cgrnro, listaEmpleados, fecHastaBucle) * porcentajeContribucion(mes, anio, tipomaq)) / ArrDatos(3, mes + 2)
                ArrDatos(18, 16) = ArrDatos(18, 16) + ArrDatos(18, mes + 2)
                'totalFila18 = totalFila18 + ArrDatos(18, mes + 2)
            End If
            

            Flog.writeline "Buscando datos - Costo pdio. Hora 50 %"
            If ArrDatos(3, mes + 2) > 0 Then
                ArrDatos(19, mes + 2) = (escalaCategoria(cgrnro, listaEmpleados, fecHastaBucle) * 1.5 * porcentajeContribucion(mes, anio, tipomaq)) / ArrDatos(3, mes + 2)
                ArrDatos(19, 16) = ArrDatos(19, 16) + ArrDatos(19, mes + 2)
                'totalFila19 = totalFila19 + ArrDatos(19, mes + 2)
            End If
            

            Flog.writeline "Buscando datos - Costo pdio. Hora 100 %"
            If ArrDatos(3, mes + 2) > 0 Then
                ArrDatos(20, mes + 2) = (escalaCategoria(cgrnro, listaEmpleados, fecHastaBucle) * 2 * porcentajeContribucion(mes, anio, tipomaq)) / ArrDatos(3, mes + 2)
                ArrDatos(20, 16) = ArrDatos(20, 16) + ArrDatos(20, mes + 2)
                'totalFila20 = totalFila20 + ArrDatos(20, mes + 2)
            End If
            
            Flog.writeline "Buscando datos - Variación en Horas"
            ArrDatos(21, mes + 2) = ArrDatos(16, mes + 2) - ArrDatos(17, mes + 2)
            ArrDatos(21, 16) = ArrDatos(21, 16) + ArrDatos(21, mes + 2)
            'totalFila21 = totalFila21 + ArrDatos(21, mes + 2)

            Flog.writeline "Buscando datos - Variación en Miles de Pesos"
            ArrDatos(22, mes + 2) = ArrDatos(21, mes + 2) * ArrDatos(18, mes + 2) * ArrDatos(2, mes + 2)
            ArrDatos(22, 16) = ArrDatos(22, 16) + ArrDatos(22, mes + 2)
            'totalFila22 = totalFila22 + ArrDatos(22, mes + 2)

            Flog.writeline "Buscando datos - Variación en Porcentajes"
            If ArrDatos(22, mes + 2) > 0 Then
                ArrDatos(23, mes + 2) = (ArrDatos(21, mes + 2) / ArrDatos(22, mes + 2)) * 100
                ArrDatos(23, 16) = ArrDatos(23, 16) + ArrDatos(23, mes + 2)
                'totalFila23 = totalFila23 + ArrDatos(23, mes + 2)
            End If
            
            Flog.writeline "Buscando datos - Horas al 50%"
            ArrDatos(24, mes + 2) = obtenerhorasSinFormula(th50, listaEmpleados, fecDesdeBucle, fecHastaBucle)
            ArrDatos(24, 16) = ArrDatos(24, 16) + ArrDatos(24, mes + 2)
            'totalFila24 = totalFila24 + ArrDatos(24, mes + 2)

            Flog.writeline "Buscando datos - Horas al 100%"
            ArrDatos(25, mes + 2) = obtenerhorasSinFormula(th100, listaEmpleados, fecDesdeBucle, fecHastaBucle)
            ArrDatos(25, 16) = ArrDatos(25, 16) + ArrDatos(25, mes + 2)
            'totalFila25 = totalFila25 + ArrDatos(25, mes + 2)

            Flog.writeline "Buscando datos - Total Horas Extras"
            ArrDatos(26, mes + 2) = ArrDatos(25, mes + 2) + ArrDatos(24, mes + 2)
            ArrDatos(26, 16) = ArrDatos(26, 16) + ArrDatos(26, mes + 2)
            'totalFila26 = totalFila26 + ArrDatos(26, mes + 2)

            Flog.writeline "Buscando datos - COSTO CON CARGAS SOCIALES"
            ArrDatos(27, mes + 2) = (ArrDatos(24, mes + 2) * ArrDatos(19, mes + 2)) + (ArrDatos(25, mes + 2) * ArrDatos(20, mes + 2))
            ArrDatos(27, 16) = ArrDatos(27, 16) + ArrDatos(27, mes + 2)
            'totalFila26 = totalFila26 + ArrDatos(26, mes + 2)

            'Actualizo el progreso
            TiempoAcumulado = GetTickCount
            StrSql = "UPDATE batch_proceso SET bprcprogreso = " & CDbl(Progreso)
            StrSql = StrSql & ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'"
            StrSql = StrSql & ", bprcempleados ='" & CStr(CEmpleadosAProc) & "' WHERE bpronro = " & NroProceso
            objConnProgreso.Execute StrSql, , adExecuteNoRecords
        Next
        
        '----------------------------------------------------------------------------------------------------------------
        If promediar > 0 Then
            Flog.writeline "Calculando promedios"
            'promedio de fila 1 - MAQUINAS PLANIFICADAS
            ArrDatos(1, 16) = ArrDatos(1, 15) / promediar
            'promedio de fila 2 - MAQUINAS ENTREGADAS A VTAS.
            ArrDatos(2, 16) = ArrDatos(2, 15) / promediar
            'promedio de fila 3 - Total de Personal de Producción
            ArrDatos(3, 16) = ArrDatos(3, 16) / promediar
            'promedio de fila 4 - Total de Personal por cantidad de Hs Reales
            ArrDatos(4, 16) = ArrDatos(4, 16) / promediar
            'promedio de fila 5 - Hs. Pdio. por Empleado
            ArrDatos(5, 16) = ArrDatos(5, 16) / promediar
            'promedio de fila 6 - Total de Horas Teoricas
            ArrDatos(6, 16) = ArrDatos(6, 16) / promediar
            'promedio de fila 7 - Total Horas Pagadas c/ Extras
            ArrDatos(7, 16) = ArrDatos(7, 16) / promediar
            'promedio de fila 8 - Total Horas Pagadas c/ Extras
            ArrDatos(8, 16) = ArrDatos(8, 15) / promediar
            'promedio de fila 9 - Horas Desviadas
            ArrDatos(9, 16) = ArrDatos(9, 16) / promediar
            'promedio de fila 10 - Feriados Pagos
            ArrDatos(10, 16) = ArrDatos(10, 16) / promediar
            'promedio de fila 11 - Licencias Pagas
            ArrDatos(11, 16) = ArrDatos(11, 16) / promediar
            'promedio de fila 12 - Enfermedad
            ArrDatos(12, 16) = ArrDatos(12, 16) / promediar
            'promedio de fila 13 - Accidente
            ArrDatos(13, 16) = ArrDatos(13, 16) / promediar
            'promedio de fila 14 - Hs Improductivas
            ArrDatos(14, 16) = ArrDatos(14, 16) / promediar
            'promedio de fila 15 - Hs Netas de Produccion
            ArrDatos(15, 16) = ArrDatos(15, 16) / promediar
            'promedio de fila 16 - Horas por Máquina Producida
            ArrDatos(16, 16) = ArrDatos(16, 16) / promediar
            'promedio de fila 17 - Standard de Fabricación
            ArrDatos(17, 16) = ArrDatos(17, 16) / promediar
            'promedio de fila 18 - Costo pdio.Gral Horas c/ Cs.Soc.
            ArrDatos(18, 16) = ArrDatos(18, 16) / promediar
            'promedio de fila 19 - Costo pdio. Hora 50 %
            ArrDatos(19, 16) = ArrDatos(19, 16) / promediar
            'promedio de fila 20 - Costo pdio. Hora 100
            ArrDatos(20, 16) = ArrDatos(20, 16) / promediar
            'promedio de fila 21 - Variación en Horas
            ArrDatos(21, 16) = ArrDatos(21, 16) / promediar
            'promedio de fila 22 - Variación en Miles de Pesos
            ArrDatos(22, 16) = ArrDatos(22, 16) / promediar
            'promedio de fila 23 - Variación en Porcentajes
            ArrDatos(23, 16) = ArrDatos(23, 16) / promediar
            'promedio de fila 24 - Horas al 50%
            ArrDatos(24, 16) = ArrDatos(24, 16) / promediar
            'promedio de fila 25 - Horas al 100%
            ArrDatos(25, 16) = ArrDatos(25, 16) / promediar
            'promedio de fila 26 - Total Horas Extras
            ArrDatos(26, 16) = ArrDatos(26, 16) / promediar
            'promedio de fila 27 - COSTO CON CARGAS SOCIALES
            ArrDatos(27, 16) = ArrDatos(27, 16) / promediar
        End If
        
        
        'inserto los datos del reporte
        Flog.writeline "Insertando datos de la cabecera, bpronro: " & NroProceso
        StrSql = " INSERT INTO gti_rep_maquinarias (bpronro, descripcion, fecgen,anio, empresa, emplogo, logoancho, logoalto, tipomaq) " & _
                 " VALUES (" & NroProceso & ",'" & historicoDesc & "'," & ConvFecha(Date) & ",'" & anio & "','" & empresa & "'" & _
                 ",'" & emplogo & "'," & logoAncho & "," & logoAlto & "," & tipomaq & ")"
        objConn.Execute StrSql, , adExecuteNoRecords
        
        Flog.writeline "Cabecera insertada, bpronro: " & NroProceso
        
        
        For fila = 1 To 27
            StrSql = " INSERT INTO gti_rep_maquinarias_det (bpronro, fila, prom1, prom2, "
            For i = 1 To 12
                StrSql = StrSql & "mes" & i & ", "
            Next
            StrSql = StrSql & "total, prom ) VALUES ( "
            
            
                StrSql = StrSql & NroProceso & ", " & fila & ", " & Replace(FormatNumber(ArrDatos(fila, 1), 2), ",", "") & ", " & Replace(FormatNumber(ArrDatos(fila, 2), 2), ",", "")
                For j = 1 To 12
                    StrSql = StrSql & ", " & Replace(FormatNumber(ArrDatos(fila, j + 2), 2), ",", "")
                Next
                StrSql = StrSql & ", " & Replace(FormatNumber(ArrDatos(fila, 15), 2), ",", "") & ", " & Replace(FormatNumber(ArrDatos(fila, 16), 2), ",", "")
            StrSql = StrSql & ") "
            objConn.Execute StrSql, , adExecuteNoRecords
        Next
        
        Flog.writeline "Insertando datos del detalle, bpronro: " & NroProceso
        
        Flog.writeline "Detalle insertado, bpronro: " & NroProceso
    Else
        Flog.writeline ""
        Flog.writeline "No existen empleados a procesar"
        Exit Sub
    End If
    

Flog.writeline ""
Flog.writeline "Proceso finalizado."


If rsEmp.State = adStateOpen Then rsEmp.Close
If rsEstruc.State = adStateOpen Then rsEstruc.Close
If rs.State = adStateOpen Then rs.Close

Set rs = Nothing

End Sub

Sub obtenerEmpleados(ByVal bpronro, ByRef listaEmpleado As String)
Dim rsEmp As New ADODB.Recordset

    StrSql = " SELECT empleado.Ternro FROM batch_empleado " & _
             " INNER JOIN empleado ON empleado.ternro =  batch_empleado.ternro " & _
             " Where bpronro = " & bpronro
    OpenRecordset StrSql, rsEmp
    listaEmpleado = "0"
    Do While Not rsEmp.EOF
        listaEmpleado = listaEmpleado & "," & rsEmp!ternro
        rsEmp.MoveNext
    Loop

If rsEmp.State = adStateOpen Then rsEmp.Close
Set rsEmp = Nothing

End Sub



Function buscarHistorico(ByVal anio As Long, ByVal fila As Long, ByVal tipomaq As Long)

Dim rsHistCab As New ADODB.Recordset
Dim rsHist As New ADODB.Recordset
    
    
    StrSql = " SELECT bpronro FROM gti_rep_maquinarias WHERE anio = " & anio & " AND tipomaq = " & tipomaq
    OpenRecordset StrSql, rsHistCab
    If Not rsHistCab.EOF Then
        'busca los datos de hsitoricos anteriores
        StrSql = " SELECT sum(prom) cant  FROM gti_rep_maquinarias_det " & _
                 " WHERE fila = " & fila & " AND bpronro = " & rsHistCab!bpronro
        OpenRecordset StrSql, rsHist
        If Not rsHist.EOF Then
            If Not IsNull(rsHist!Cant) Then
                buscarHistorico = rsHist!Cant
            Else
                buscarHistorico = 0
            End If
        Else
            buscarHistorico = 0
        End If
    Else
        buscarHistorico = 0
    End If
    
If rsHist.State = adStateOpen Then rsHist.Close
If rsHistCab.State = adStateOpen Then rsHistCab.Close

Set rsHistCab = Nothing
Set rsHist = Nothing

End Function

Function maquinasPlanificadas(ByVal mes As Long, ByVal anio As Long, ByVal tipomaq As Long)
Dim rsMaq As New ADODB.Recordset

    StrSql = " SELECT maqplanif FROM gti_maquinaria " & _
             " WHERE mes = " & mes & " AND anio = " & anio & " AND tipomaq = " & tipomaq
    OpenRecordset StrSql, rsMaq
    If Not rsMaq.EOF Then
        maquinasPlanificadas = rsMaq!maqplanif
    Else
        maquinasPlanificadas = 0
    End If
    
If rsMaq.State = adStateOpen Then rsMaq.Close
Set rsMaq = Nothing

End Function

Function HsStandard(ByVal mes As Long, ByVal anio As Long, ByVal tipomaq As Long)
Dim rsMaq As New ADODB.Recordset

    StrSql = " SELECT maqhs FROM gti_maquinaria " & _
             " WHERE mes = " & mes & " AND anio = " & anio & " AND tipomaq = " & tipomaq
    OpenRecordset StrSql, rsMaq
    If Not rsMaq.EOF Then
        HsStandard = rsMaq!maqhs
    Else
        HsStandard = 0
    End If
    
If rsMaq.State = adStateOpen Then rsMaq.Close
Set rsMaq = Nothing

End Function


Function maquinasEntregadas(ByVal mes As Long, ByVal anio As Long, ByVal tipomaq As Long)
Dim rsMaq As New ADODB.Recordset

    StrSql = " SELECT maqentreg FROM gti_maquinaria " & _
             " WHERE mes = " & mes & " AND anio = " & anio & " AND tipomaq = " & tipomaq
    OpenRecordset StrSql, rsMaq
    If Not rsMaq.EOF Then
        maquinasEntregadas = rsMaq!maqentreg
    Else
        maquinasEntregadas = 0
    End If
    
If rsMaq.State = adStateOpen Then rsMaq.Close
Set rsMaq = Nothing
End Function

Function empleadosEstructura(ByVal estrnro1 As String, ByVal estrnro2 As String, ByVal fecdesde As String, ByVal fechasta As String)
Dim rsEst As New ADODB.Recordset

    StrSql = " SELECT count(distinct empleado.ternro) cant FROM empleado " & _
            " INNER JOIN his_estructura he1 ON he1.ternro = empleado.ternro AND he1.estrnro = " & estrnro1 & _
            " INNER JOIN his_estructura he2 ON he2.ternro = empleado.ternro AND he2.estrnro = " & estrnro2 & _
            " WHERE ((he1.htetdesde <= " & ConvFecha(fecdesde) & " AND (he1.htethasta is null or he1.htethasta >= " & ConvFecha(fechasta) & _
            " or he1.htethasta >= " & ConvFecha(fecdesde) & ")) OR(he1.htetdesde >= " & ConvFecha(fecdesde) & " AND (he1.htetdesde <= " & ConvFecha(fechasta) & "))) " & _
            " AND ((he2.htetdesde <= " & ConvFecha(fecdesde) & " AND (he2.htethasta is null or he2.htethasta >= " & ConvFecha(fechasta) & _
            " or he2.htethasta >= " & ConvFecha(fecdesde) & ")) OR(he2.htetdesde >= " & ConvFecha(fecdesde) & " AND (he2.htetdesde <= " & ConvFecha(fechasta) & "))) " & _
            " AND empest = -1 "
    OpenRecordset StrSql, rsEst
    If Not rsEst.EOF Then
        If Not IsNull(rsEst!Cant) Then
            empleadosEstructura = rsEst!Cant
        Else
            empleadosEstructura = 0
        End If
    Else
        empleadosEstructura = 0
    End If


If rsEst.State = adStateOpen Then rsEst.Close
Set rsEst = Nothing

End Function

Function obtenerhoras(ByVal thnormal As String, ByVal empleados As String, ByVal fecdesde As String, ByVal fechasta As String, ByVal hsdiasHab As Double)
Dim rsHs As New ADODB.Recordset
Dim salida As Double


    StrSql = " SELECT SUM(adcanthoras) cant FROM gti_acumdiario WHERE thnro in (" & thnormal & ")" & _
             " and adfecha >= " & ConvFecha(fecdesde) & " AND adfecha <= " & ConvFecha(fechasta) & " AND ternro in (" & empleados & ") "
    OpenRecordset StrSql, rsHs
    If Not rsHs.EOF Then
        If Not IsNull(rsHs!Cant) Then
            salida = rsHs!Cant
        Else
            salida = 0
        End If
    Else
        salida = 0
    End If

    obtenerhoras = Round(salida / (diashabiles(fecdesde, fechasta) * hsdiasHab))
If rsHs.State = adStateOpen Then rsHs.Close
Set rsHs = Nothing
             
End Function

Function diashabiles(ByVal fecdesde As String, ByVal fechasta As String) As Long
Dim salida As Long
Dim fecha As Date

    salida = 0
    'hasta = DateDiff("d", fecdesde, fechasta) + 1
    For fecha = fecdesde To fechasta
        If Not esferiado(fecha, 3) Then
            If Weekday(fecha) <> 1 And Weekday(fecha) <> 7 Then
                salida = salida + 1
            End If
        End If
    Next
    
    diashabiles = salida
End Function

Function esferiado(ByVal fecha As Date, ByVal pais As Long) As Boolean
Dim rsFeriado As New ADODB.Recordset

    StrSql = "select ferinro from feriado WHERE fericodext = " & pais & " AND ferifecha = " & ConvFecha(fecha)
    OpenRecordset StrSql, rsFeriado
    If Not rsFeriado.EOF Then
        esferiado = True
    Else
        esferiado = False
    End If
If rsFeriado.State = adStateOpen Then rsFeriado.Close
Set rsFeriado = Nothing
End Function

Function obtenerhorasSinFormula(ByVal thnormal As String, ByVal empleados As String, ByVal fecdesde As String, ByVal fechasta As String) As Double
Dim rsHs As New ADODB.Recordset
Dim salida As Double


    StrSql = " SELECT SUM(adcanthoras) cant FROM gti_acumdiario WHERE thnro in (" & thnormal & ")" & _
             " and adfecha >= " & ConvFecha(fecdesde) & " AND adfecha <= " & ConvFecha(fechasta) & " AND ternro in (" & empleados & ") "
    OpenRecordset StrSql, rsHs
    If Not rsHs.EOF Then
        If Not IsNull(rsHs!Cant) Then
            salida = rsHs!Cant
        Else
            salida = 0
        End If
    Else
        salida = 0
    End If

    obtenerhorasSinFormula = salida
If rsHs.State = adStateOpen Then rsHs.Close
Set rsHs = Nothing
             
End Function

Function escalaCategoria(ByVal cgrnro As String, ByVal listaEmpleados As String, ByVal fecHastaBucle As String) As Double
Dim rsEstructura As New ADODB.Recordset
Dim rsEscala As New ADODB.Recordset
Dim salida As Double
Dim ternro As Long
'Dim cantEmpleados As Long
Dim i As Long
Dim arrEmpleado
    
    salida = 0
    'cantEmpleados = 0
    arrEmpleado = Split(listaEmpleados, ",")
    For i = 0 To UBound(arrEmpleado)
        ternro = arrEmpleado(i)
        StrSql = " SELECT convenio.estrnro conv, categoria.estrnro cat FROM empleado " & _
                 " INNER JOIN his_estructura convenio ON convenio.ternro = empleado.ternro AND convenio.tenro = 19 " & _
                 " INNER JOIN his_estructura categoria ON categoria.ternro = empleado.ternro AND categoria.tenro = 3 " & _
                 " WHERE empleado.ternro = " & ternro & " AND (convenio.htetdesde <= " & ConvFecha(fecHastaBucle) & " AND (convenio.htethasta >= " & ConvFecha(fecHastaBucle) & " OR convenio.htethasta is null)) " & _
                 " AND (categoria.htetdesde <= " & ConvFecha(fecHastaBucle) & " AND (categoria.htethasta >= " & ConvFecha(fecHastaBucle) & " OR categoria.htethasta is null)) "
        OpenRecordset StrSql, rsEstructura
        If Not rsEstructura.EOF Then
            'cantEmpleados = cantEmpleados + 1
            If Not IsNull(rsEstructura!conv) And Not IsNull(rsEstructura!conv) Then
                StrSql = " SELECT vgrvalor FROM valgrilla " & _
                         " WHERE cgrnro = " & cgrnro & " And vgrcoor_1 = " & rsEstructura!conv & _
                         " And vgrcoor_2 = " & rsEstructura!cat & " And vgrorden = 3 "
                OpenRecordset StrSql, rsEscala
                If Not rsEscala.EOF Then
                    salida = salida + rsEscala!vgrvalor
                End If
            End If
        End If
    Next

    escalaCategoria = salida '* cantEmpleados

If rsEstructura.State = adStateOpen Then rsEstructura.Close
Set rsEstructura = Nothing

If rsEscala.State = adStateOpen Then rsEscala.Close
Set rsEscala = Nothing

End Function

Function porcentajeContribucion(ByVal mes As Integer, ByVal anio As String, ByVal tipomaq As String) As Double
Dim rsPorcentaje As New ADODB.Recordset

    StrSql = " SELECT maqporc FROM gti_maquinaria " & _
             " WHERE mes = " & mes & " AND anio = " & anio & " AND tipomaq = " & tipomaq
    OpenRecordset StrSql, rsPorcentaje
    If Not rsPorcentaje.EOF Then
        porcentajeContribucion = 1 + (rsPorcentaje!maqporc / 100)
    Else
        porcentajeContribucion = 1
    End If
    
If rsPorcentaje.State = adStateOpen Then rsPorcentaje.Close
Set rsPorcentaje = Nothing

End Function
