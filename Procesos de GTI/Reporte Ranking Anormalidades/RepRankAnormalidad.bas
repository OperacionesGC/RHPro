Attribute VB_Name = "mdlRepRankAnormalidad"
Option Explicit

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Const Version = "1.00" 'Version inicial
Const FechaVersion = "30/07/2013"
'Deluchi Ezequiel - CAS - 19692 - Reporte de ranking de amonestaciones (licencias activas y novedades horarias)
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
    Archivo = PathFLog & "RankingAnormalidades_" & CStr(NroProceso) & ".log"

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
        If Not IsNull(rs!bprcparam) Then
            If Len(rs!bprcparam) >= 1 Then
                Flog.writeline "Inicio de Reporte Ranking de Anormalidades: " & " " & Now
                Call Rep_ranking_anormalidades(rs!bprcparam, NroProceso)
            End If
        End If
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
    Flog.writeline Espacios(Tabulador * 0) & "Fin de Reporte de Ranking de Anormalidades: " & " " & Now
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
Private Sub Rep_ranking_anormalidades(parametros As String, NroProceso As Long)

Dim rs As New ADODB.Recordset
Dim i As Long
Dim rsEmp As New ADODB.Recordset
Dim rsEstruc As New ADODB.Recordset
Dim historicoDesc As String
Dim tenro1 As Long
Dim tenro2 As Long
Dim tenro3 As Long
Dim estrnro1 As Long
Dim estrnro2 As Long
Dim estrnro3 As Long

Dim tenro1Empleado As Long
Dim tenro2Empleado As Long
Dim tenro3Empleado As Long
Dim estrnro1Empleado As Long
Dim estrnro2Empleado As Long
Dim estrnro3Empleado As Long

Dim ArrParametros
Dim Ternro
Dim fecDesde
Dim fecHasta
Dim estrnroEmpresa
Dim crearCabecera As Boolean
Dim puntajeLicencias As Double
Dim puntajeNovedades As Double
Dim rannro As Double
Dim nombre As String
Dim listaLicencias As String
Dim listaNovedades As String
Dim puntajeTotal As Double
Dim sector As String
Dim descLicencia As String
Dim descNovedad As String

'======================================================================================================
'SE VALIDA Y LEVANTAN PARAMETROS.
'======================================================================================================
        
        '0@0@0@0@0@0@01/05/2013@20/05/2013@1240
    ArrParametros = Split(parametros, "@")

    tenro1 = ArrParametros(0)
    estrnro1 = ArrParametros(1)
    tenro2 = ArrParametros(2)
    estrnro2 = ArrParametros(3)
    tenro3 = ArrParametros(4)
    estrnro3 = ArrParametros(5)
    fecDesde = ArrParametros(6)
    fecHasta = ArrParametros(7)
    estrnroEmpresa = ArrParametros(8)
    
    'levanto la lista de licencias y de novedades configuradas en el confrep 408
    StrSql = " SELECT confnrocol, confval FROM confrep WHERE repnro = 408 "
    OpenRecordset StrSql, rs
    listaLicencias = "0"
    listaNovedades = "0"
    Do While Not rs.EOF
        Select Case CLng(rs!confnrocol)
            Case 1: 'licencias
                listaLicencias = listaLicencias & "," & rs!confval
            Case 2: 'Novedades
                listaNovedades = listaNovedades & "," & rs!confval
        End Select
        rs.MoveNext
    Loop
    
    historicoDesc = NroProceso & " - Ranking Anormalidades para el " & fecDesde & " hasta " & fecHasta
  
    Flog.writeline ""
    Flog.writeline "Parametros:" & parametros
    Flog.writeline ""

    Call obtenerEmpleados(rsEmp, NroProceso)
        
    If Not rsEmp.EOF Then
        CEmpleadosAProc = rsEmp.RecordCount
        If CEmpleadosAProc = 0 Then
            CEmpleadosAProc = 1
        End If
    End If
    
    If Not rsEmp.EOF Then
        crearCabecera = False
        Progreso = 0
        Flog.writeline "Cantidad de empleados a procesar: " & CEmpleadosAProc
        IncPorc = (100 / CEmpleadosAProc)
    
        TiempoInicialProceso = GetTickCount
        i = 0
    
        Do While Not rsEmp.EOF
            'Actualizo el progreso
            Progreso = Progreso + IncPorc
            Ternro = rsEmp!Ternro
            Flog.writeline "Analizando el empleado: " & Ternro
                          
            nombre = rsEmp!terape & IIf(IsNull(rsEmp!terape2), ", ", " " & rsEmp!terape2 & ", ")
            nombre = nombre & rsEmp!ternom & IIf(IsNull(rsEmp!ternom2), "", " " & rsEmp!ternom2)
            puntajeTotal = 0
            'busco estructuras del empleado segun filtro
            If tenro1 <> 0 Then
                StrSql = " SELECT he1.tenro tenro1, he1.estrnro estrnro1 "
            End If
            If tenro2 <> 0 Then
                StrSql = StrSql & " ,he2.tenro tenro2, he2.estrnro estrnro2 "
            End If
            If tenro3 <> 0 Then
                StrSql = StrSql & " ,he3.tenro tenro3 ,he3.estrnro estrnro3 "
            End If
            If tenro1 <> 0 Then
                StrSql = StrSql & " FROM empleado " & _
                         " INNER JOIN his_estructura he1 ON he1.ternro = empleado.ternro AND he1.tenro = " & tenro1 & _
                         " AND ((he1.htetdesde <= " & ConvFecha(fecDesde) & " AND (he1.htethasta is null or he1.htethasta >= " & ConvFecha(fecHasta) & _
                         " OR he1.htethasta >= " & ConvFecha(fecDesde) & ")) OR (he1.htetdesde >= " & ConvFecha(fecDesde) & " AND (he1.htetdesde <= " & ConvFecha(fecHasta) & "))) "
                
                If tenro2 <> 0 Then
                StrSql = StrSql & " INNER JOIN his_estructura he2 ON he2.ternro = empleado.ternro AND he2.tenro = " & tenro2 & _
                             " AND ((he2.htetdesde <= " & ConvFecha(fecDesde) & " AND (he2.htethasta is null or he2.htethasta >= " & ConvFecha(fecHasta) & _
                             " OR he2.htethasta >= " & ConvFecha(fecDesde) & ")) OR (he2.htetdesde >= " & ConvFecha(fecDesde) & " AND (he2.htetdesde <= " & ConvFecha(fecHasta) & "))) "
                End If
                    
                If tenro3 <> 0 Then
                StrSql = StrSql & " INNER JOIN his_estructura he3 ON he3.ternro = empleado.ternro AND he3.tenro = " & tenro3 & _
                             " AND ((he3.htetdesde <= " & ConvFecha(fecDesde) & " AND (he3.htethasta is null or he3.htethasta >= " & ConvFecha(fecHasta) & _
                             " OR he3.htethasta >= " & ConvFecha(fecDesde) & ")) OR (he3.htetdesde >= " & ConvFecha(fecDesde) & " AND (he3.htetdesde <= " & ConvFecha(fecHasta) & "))) "
                End If
                StrSql = StrSql & " WHERE empleado.ternro = " & rsEmp!Ternro
                If estrnro1 <> 0 Then
                    StrSql = StrSql & " AND he1.estrnro = " & estrnro1
                End If
                If estrnro2 <> 0 Then
                    StrSql = StrSql & " AND he2.estrnro = " & estrnro2
                End If
                If estrnro3 <> 0 Then
                    StrSql = StrSql & " AND he3.estrnro = " & estrnro3
                End If
                
                OpenRecordset StrSql, rsEstruc
                If Not rsEstruc.EOF Then
                    tenro1Empleado = 0
                    tenro2Empleado = 0
                    tenro3Empleado = 0
                    estrnro1Empleado = 0
                    estrnro2Empleado = 0
                    estrnro3Empleado = 0
                    If tenro1 <> 0 Then
                        tenro1Empleado = rsEstruc!tenro1
                        estrnro1Empleado = rsEstruc!estrnro1
                    End If
                    If tenro2 <> 0 Then
                        tenro2Empleado = rsEstruc!tenro2
                        estrnro2Empleado = rsEstruc!estrnro2
                    End If
                    If tenro3 <> 0 Then
                        tenro3Empleado = rsEstruc!tenro3
                        estrnro3Empleado = rsEstruc!estrnro3
                    End If
                End If
            End If
            'Fin busqueda estructuras del empleado
            
            'obtengo el sector del empleado
            StrSql = " SELECT estrdabr FROM his_estructura " & _
                     " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro AND his_estructura.tenro = 2 " & _
                     " WHERE  his_estructura.ternro = " & Ternro & _
                     " AND ((his_estructura.htetdesde <= " & ConvFecha(fecDesde) & " AND (his_estructura.htethasta is null or his_estructura.htethasta >= " & ConvFecha(fecHasta) & _
                     " or his_estructura.htethasta >= " & ConvFecha(fecDesde) & ")) OR (his_estructura.htetdesde >= " & ConvFecha(fecDesde) & " AND (his_estructura.htetdesde <= " & ConvFecha(fecHasta) & "))) "

            OpenRecordset StrSql, rs
            If Not rs.EOF Then
                sector = rs!estrdabr
            Else
                sector = ""
            End If
            
            'Busco las licencias del empleado para analizarlas
            StrSql = " SELECT emp_licnro, elfechadesde, elfechahasta, elcantdias, emp_lic.tdnro, tddesc FROM emp_lic " & _
                     " INNER JOIN tipdia ON tipdia.tdnro = emp_lic.tdnro AND tdest = -1 " & _
                     " WHERE  empleado = " & Ternro & " AND licestnro = 2 " & _
                     " AND ((elfechadesde <= " & ConvFecha(fecDesde) & " AND (elfechahasta >= " & ConvFecha(fecHasta) & " or elfechahasta >= " & ConvFecha(fecDesde) & ")) " & _
                     " OR (elfechadesde >= " & ConvFecha(fecDesde) & " AND (elfechadesde <= " & ConvFecha(fecHasta) & "))) " & _
                     " AND emp_lic.tdnro in ( " & listaLicencias & ") AND eldiacompleto = -1 ORDER BY emp_lic.tdnro "
            OpenRecordset StrSql, rs
            
            If Not rs.EOF Then
                Do While Not rs.EOF
                    puntajeLicencias = 0
                    
                    'analizo la licencia y devuelvo el puntaje obtenido para el empleado
                    Flog.writeline "Analizando la licencia: " & rs!emp_licnro
                    Call analizarLicencia(rs, fecDesde, fecHasta, puntajeLicencias)
                    
                    'Controlo que halla al menos un datos para crear la cabecera
                    If Not crearCabecera And (puntajeLicencias > 0) Then
                        'con esta asignacion me aseguro que no se vuelvan a crear cabeceras para el mismo reporte
                        crearCabecera = True
                        StrSql = " INSERT INTO rep_rank_anormalidad (radesc,bpronro,rafecha,rahora,tenro1,estrnro1,tenro2,estrnro2,tenro3,estrnro3) VALUES " & _
                                 " ('" & historicoDesc & "'," & NroProceso & "," & ConvFecha(Date) & ",'" & Left(Time, 8) & "'," & _
                                  tenro1 & "," & estrnro1 & "," & tenro2 & "," & estrnro2 & "," & tenro3 & "," & estrnro3 & ")"
                        objConn.Execute StrSql, , adExecuteNoRecords
                        rannro = getLastIdentity(objConn, "rep_lic_rank")
                        Flog.writeline "Creada la cabecera para el reporte: " & NroProceso
                    End If
                    
                    If puntajeLicencias > 0 Then
                        'inserto el detalle del empleado si hay puntaje
                        StrSql = " INSERT INTO rep_rank_anormalidad_det (ranro,ternro,ratipoorigen,ratipocodorigen,racodorigen,empleg,nombre,puntaje,sector,anordesc,tenro1,estrnro1,tenro2,estrnro2,tenro3,estrnro3) VALUES " & _
                                 " (" & rannro & "," & rsEmp!Ternro & ",1," & rs!tdnro & "," & rs!emp_licnro & "," & rsEmp!empleg & ",'" & nombre & "'," & puntajeLicencias & " ,'" & sector & "','" & rs!tddesc & "'," & tenro1Empleado & " ," & estrnro1Empleado & _
                                 " ," & tenro2Empleado & " ," & estrnro2Empleado & " ," & tenro3Empleado & " ," & estrnro3Empleado & ")"
                        objConn.Execute StrSql, , adExecuteNoRecords
                        puntajeTotal = puntajeTotal + puntajeLicencias
                        Flog.writeline "Insertado el detalle para el empleado: " & rsEmp!Ternro & ", cabecera de reporte: " & rannro & " y proceso: " & NroProceso
                    Else
                        Flog.writeline "No hay puntaje de licencias para el empleado: " & rsEmp!Ternro
                    End If

                    rs.MoveNext
                
                Loop
            Else 'else de consulta de licencias para el empleado
                Flog.writeline "No existen licencias a procesar para el empleado: " & rsEmp!Ternro & " en el rango de fechas " & fecDesde & " - " & fecHasta
            End If

            'Busco las novedades del empleado para analizarlas
            StrSql = " SELECT gtnovnro, gnovnro, gnovdesde, gnovhasta, gnovdesabr FROM gti_novedad " & _
                     " WHERE ((gnovdesde <= " & ConvFecha(fecDesde) & " AND (gnovhasta >= " & ConvFecha(fecHasta) & " or gnovhasta >= " & ConvFecha(fecDesde) & ")) " & _
                     " OR (gnovdesde >= " & ConvFecha(fecDesde) & " AND (gnovdesde <= " & ConvFecha(fecHasta) & "))) AND gnovotoa = " & rsEmp!Ternro & _
                     " AND gtnovnro in (" & listaNovedades & ") ORDER BY gnovnro "
            OpenRecordset StrSql, rs
            If Not rs.EOF Then
                '-----------------------------------------------------------
                Do While Not rs.EOF
                    
                    puntajeNovedades = 0
                    'analizo la licencia y devuelvo el puntaje obtenido para el empleado
                    Flog.writeline "Analizando la novedad horaria: " & rs!gnovnro
                    Call analizarNovedad(rs, fecDesde, fecHasta, puntajeNovedades)
                    
                    'Controlo que halla al menos un datos para crear la cabecera
                    If Not crearCabecera And (puntajeNovedades > 0) Then
                        'con esta asignacion me aseguro que no se vuelvan a crear cabeceras para el mismo reporte
                        crearCabecera = True
                        StrSql = " INSERT INTO rep_rank_anormalidad (radesc,bpronro,rafecha,rahora,tenro1,estrnro1,tenro2,estrnro2,tenro3,estrnro3) VALUES " & _
                                 " ('" & historicoDesc & "'," & NroProceso & "," & ConvFecha(Date) & ",'" & Left(Time, 8) & "'," & _
                                  tenro1 & "," & estrnro1 & "," & tenro2 & "," & estrnro2 & "," & tenro3 & "," & estrnro3 & ")"
                        objConn.Execute StrSql, , adExecuteNoRecords
                        rannro = getLastIdentity(objConn, "rep_lic_rank")
                        Flog.writeline "Creada la cabecera para el reporte: " & NroProceso
                    End If
                    
                    If puntajeNovedades > 0 Then
                        'inserto el detalle del empleado si hay puntaje
                        StrSql = " INSERT INTO rep_rank_anormalidad_det (ranro,ternro,ratipoorigen,ratipocodorigen,racodorigen,empleg,nombre,puntaje,sector,anordesc,tenro1,estrnro1,tenro2,estrnro2,tenro3,estrnro3) VALUES " & _
                                 " (" & rannro & "," & rsEmp!Ternro & ",2," & rs!gtnovnro & "," & rs!gnovnro & "," & rsEmp!empleg & ",'" & nombre & "'," & puntajeNovedades & " ,'" & sector & "','" & rs!gnovdesabr & "'," & tenro1Empleado & " ," & estrnro1Empleado & _
                                 " ," & tenro2Empleado & " ," & estrnro2Empleado & " ," & tenro3Empleado & " ," & estrnro3Empleado & ")"
                        objConn.Execute StrSql, , adExecuteNoRecords
                        puntajeTotal = puntajeTotal + puntajeNovedades
                        Flog.writeline "Insertado el detalle para el empleado: " & rsEmp!Ternro & ", cabecera de reporte: " & rannro & " y proceso: " & NroProceso
                    Else
                        Flog.writeline "No hay puntaje de novedades para el empleado: " & rsEmp!Ternro
                    End If

                    rs.MoveNext
                
                Loop
            Else
                Flog.writeline "No existen novedades horarias a procesar para el empleado: " & rsEmp!Ternro & " en el rango de fechas " & fecDesde & " - " & fecHasta
            End If
            
            'actualizalo el puntaje total para el empleado
            StrSql = " update rep_rank_anormalidad_det set " & _
                     " rapuntajetotal = " & puntajeTotal & _
                     " Where ranro = " & rannro & " And Ternro = " & rsEmp!Ternro
            objConn.Execute StrSql, , adExecuteNoRecords
            
            'Actualizo el progreso
            TiempoAcumulado = GetTickCount
            StrSql = "UPDATE batch_proceso SET bprcprogreso = " & CDbl(Progreso)
            StrSql = StrSql & ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'"
            StrSql = StrSql & ", bprcempleados ='" & CStr(CEmpleadosAProc) & "' WHERE bpronro = " & NroProceso
            objConn.Execute StrSql, , adExecuteNoRecords
            
            rsEmp.MoveNext
        Loop
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

Sub analizarLicencia(ByRef rs As Recordset, ByVal fecDesde, ByVal fecHasta, ByRef puntajeLicencias As Double)
Dim Dias As Integer
Dim fd As Date
Dim fh As Date
    
    Dias = 0
    
    If CDate(fecDesde) > CDate(rs!elfechadesde) Then
        fd = CDate(fecDesde)
    Else
        fd = CDate(rs!elfechadesde)
    End If
    
    If CDate(fecHasta) < rs!elfechahasta Then
        fh = CDate(fecHasta)
    Else
        fh = rs!elfechahasta
    End If
    
    Dias = DateDiff("d", fd, fh) + 1

    puntajeLicencias = puntajeLicencias + Dias
        
End Sub

Sub analizarNovedad(ByRef rs As Recordset, ByVal fecDesde, ByVal fecHasta, ByRef puntajeNovedades As Double)
Dim Dias As Integer
Dim fd As Date
Dim fh As Date
    
    Dias = 0
    
    If CDate(fecDesde) > CDate(rs!gnovdesde) Then
        fd = CDate(fecDesde)
    Else
        fd = CDate(rs!gnovdesde)
    End If
    
    If CDate(fecHasta) < rs!gnovhasta Then
        fh = CDate(fecHasta)
    Else
        fh = rs!gnovhasta
    End If
    
    Dias = DateDiff("d", fd, fh) + 1

    puntajeNovedades = puntajeNovedades + Dias
        
End Sub



Sub obtenerEmpleados(ByRef rsEmp As Recordset, ByVal bpronro)
    StrSql = " SELECT * FROM batch_empleado " & _
             " INNER JOIN empleado ON empleado.ternro =  batch_empleado.ternro " & _
             " Where bpronro = " & bpronro
    OpenRecordset StrSql, rsEmp
End Sub



Function descripcionEstructura(ByVal Tenro As Long, ByVal estrnro As Long)
'BUSCA DESCRIPCION DE ESTRUCTURA
    Dim StrSql
    Dim rs2 As New ADODB.Recordset
    
    StrSql = "SELECT tedabr, estrdabr FROM tipoestructura INNER JOIN estructura ON tipoestructura.tenro = estructura.tenro "
    StrSql = StrSql & "WHERE tipoestructura.tenro = " & Tenro
    If estrnro <> 0 Then
        StrSql = StrSql & " AND estrnro = " & estrnro
    End If

    rs2.Open StrSql, objConn
    If Not rs2.EOF Then
        descripcionEstructura = rs2!tedabr & ": " & rs2!estrdabr
    End If
    rs2.Close
    
    
End Function

