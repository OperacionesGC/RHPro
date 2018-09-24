Attribute VB_Name = "mdlReplicRank"
Option Explicit

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Const Version = "1.00"
'Const FechaVersion = "22/05/2013"
'Deluchi Ezequiel - CAS-18894

'Const Version = "1.01"
'Const FechaVersion = "19/06/2013"
'Deluchi Ezequiel - CAS-18894 - Se agrego columnas de estructuras al detalle para cada empleado.


'Const Version = "1.02"
'Const FechaVersion = "15/07/2013"
'15/07/2013 - Maurico Zwenger - CAS-18894 - Cambio en forma de calculo de valor para ranking segun dias corridos teniendo en cuenta
'           fecha ingresada en filtro, fechas de inicio y fin de licencia, "A partir del dia" y "Valor" configurado en el ranking
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Const Version = "1.03"
Const FechaVersion = "23/08/2013"
'23/08/2013 - Maurico Zwenger - CAS-18894 - Se corrigio calculo de dias a tener en cuenta, al hacer diferencia entre fechas quedaba fuera el dia de inicio
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

'Dim objconnMain As New ADODB.Connection
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
    'Archivo = PathFLog & "RepAusentismo-" & CStr(NroProceso) & Format(Now, "DD-MM-YYYY") & ".log"
    Archivo = PathFLog & "RepLicRank_" & CStr(NroProceso) & ".log"

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
        'MyBeginTrans
            StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcprogreso = 0, bprcestado = 'Error de Version', bprcpid = " & PID & " WHERE bpronro = " & NroProceso
            objConnProgreso.Execute StrSql, , adExecuteNoRecords
        'MyCommitTrans
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
    'rs.Open StrSql, objConn
    OpenRecordset StrSql, rs
    If Not rs.EOF Then
        'FDesde = rs!bprcfecdesde
        'FHasta = rs!bprcfechasta
        'l_iduser = rs!iduser
        If Not IsNull(rs!bprcparam) Then
            If Len(rs!bprcparam) >= 1 Then
                'pos = InStr(1, rs!bprcparam, ",")
                'NroReporte = CLng(Left(rs!bprcparam, pos - 1))
                'StrParametros = Right(rs!bprcparam, Len(rs!bprcparam) - (pos))
                'If rs.State = adStateOpen Then rs.Close
                Flog.writeline "Inicio de Reporte Licencias Rankedas: " & " " & Now
                Call Rep_licencias_rankeadas(rs!bprcparam, NroProceso)
            End If
        End If
    Else
        Exit Sub
    End If
    
   
    
    'Call Rep_AcumD(NroReporte, NroProceso, FDesde, FHasta, StrParametros)
    
    
    ' poner el bprcestado en procesado
    StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado', bprcprogreso = 100 WHERE bpronro = " & NroProceso
    objConn.Execute StrSql, , adExecuteNoRecords

    If objConn.State = adStateOpen Then objConn.Close
    If objConnProgreso.State = adStateOpen Then objConnProgreso.Close
    
    Set rs_Batch_Proceso = Nothing
    Set rs_His_Batch_Proceso = Nothing

Final:
    Flog.writeline Espacios(Tabulador * 0) & "Fin de Reporte de Licencias Rankeadas: " & " " & Now
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 0) & "---------------------------------------------------------------------------------"
    'Flog.writeline "Cantidad de Lecturas en BD          : " & Cantidad_de_OpenRecordset
    'Flog.writeline "Cantidad de llamadas a politicas    : " & Cantidad_Call_Politicas
    Flog.writeline Espacios(Tabulador * 0) & "---------------------------------------------------------------------------------"
    Flog.Close
Exit Sub

ce:
    Flog.writeline Espacios(Tabulador * 0) & "Reporte abortado por Error:" & " " & Now
    Flog.writeline Espacios(Tabulador * 0) & "Reporte abortado por :" & Err.Description
    Flog.writeline Espacios(Tabulador * 0) & "Ultimo SQL " & StrSql
    'MyRollbackTrans
    StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Error' WHERE bpronro = " & NroProceso
    objConn.Execute StrSql, , adExecuteNoRecords
End Sub
Private Sub Rep_licencias_rankeadas(parametros As String, NroProceso As Long)

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
Dim ArrParametros
Dim Ternro
Dim fecDesde
Dim fecHasta
Dim estrnroEmpresa
Dim crearCabecera As Boolean
Dim sumaRanking As Double
Dim licrannro As Double
Dim nombre As String

'======================================================================================================
'SE VALIDA Y LEVANTAN PARAMETROS.
'======================================================================================================
'If Not IsNull(parametros) And Len(parametros) >= 1 Then
    'If Len(parametros) >= 1 Then
        
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
    
    historicoDesc = NroProceso & " - Licencias Rankeadas para el " & fecDesde & " hasta " & fecHasta
  
    Flog.writeline ""
    Flog.writeline "Parametros:" & parametros
    Flog.writeline ""

    Call obtenerEmpleados(rsEmp, NroProceso)
        
    If Not rsEmp.EOF Then
        totalEmpleados = rsEmp.RecordCount
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
            sumaRanking = 0
            'CAMBIO DE EMPLEADO Y GUARDO REGISTRO CON LOS TOTALES DE HORAS.
            Progreso = Progreso + IncPorc
            Flog.writeline "Analizando el empleado: " & Ternro
            Ternro = rsEmp!Ternro
               
            nombre = rsEmp!terape & IIf(IsNull(rsEmp!terape2), ", ", " " & rsEmp!terape2 & ", ")
            nombre = nombre & rsEmp!ternom & IIf(IsNull(rsEmp!ternom2), "", " " & rsEmp!ternom2)
            
            'Busco las licencias del empleado para analizarlas
            
            
'            StrSql = " SELECT emp_licnro, elfechadesde, elfechahasta, elcantdias,valor, dias FROM emp_lic " & _
'                     " INNER JOIN tipdia_rank ON tipdia_rank.tdnro = emp_lic.tdnro " & _
'                     " WHERE  empleado = " & Ternro & " AND licestnro = 2 " & _
'                     " AND ((elfechadesde <= " & ConvFecha(fecDesde) & " AND (elfechahasta >= " & ConvFecha(fecHasta) & " or elfechahasta >= " & ConvFecha(fecDesde) & ")) " & _
'                     " OR (elfechadesde >= " & ConvFecha(fecDesde) & " AND (elfechadesde <= " & ConvFecha(fecHasta) & "))) " & _
'                     " ORDER BY emp_lic.tdnro "
            
            '15/07/2013 - MDZ - CAS-18894
            StrSql = " SELECT emp_licnro, elfechadesde, elfechahasta, elcantdias,valor, dias FROM emp_lic " & _
                     " INNER JOIN tipdia_rank ON tipdia_rank.tdnro = emp_lic.tdnro " & _
                     " WHERE  empleado = " & Ternro & " AND licestnro = 2 AND (" & _
                     " elfechadesde BETWEEN " & ConvFecha(fecDesde) & " AND " & ConvFecha(fecHasta) & " OR" & _
                     " elfechahasta BETWEEN " & ConvFecha(fecDesde) & " AND " & ConvFecha(fecHasta) & " OR" & _
                     " (elfechadesde< " & ConvFecha(fecDesde) & " AND elfechahasta>" & ConvFecha(fecHasta) & "))" & _
                     " ORDER BY emp_lic.tdnro "

            OpenRecordset StrSql, rs
            
            If Not rs.EOF Then

                Do While Not rs.EOF
                    'analizo la licencia y devuelvo el puntaje obtenido para el empleado
                    Flog.writeline "Analizando la licencia: " & rs!emp_licnro
                    
                    
                    'Call analizarLicencia(rs, fecDesde, fecHasta, sumaRanking)
                    
                    '15/07/2013 - MDZ - CAS-18894 - calculo de valor para ranking segun dias corridos teniendo en cuenta rango fechas ingresada en filtro,
                    '                   fechas de inicio y fin de licencia y "A partir del dia" configurado en el ranking
                    Dim Dias As Integer
                    Dias = 0
                    Dim fd As Date
                    Dim fh As Date
                    
                    If CDate(fecDesde) > DateAdd("d", rs!Dias, rs!elfechadesde) Then
                        fd = CDate(fecDesde)
                    Else
                        fd = DateAdd("d", rs!Dias, rs!elfechadesde)
                    End If
                    
                    If CDate(fecHasta) < rs!elfechahasta Then
                        fh = CDate(fecHasta)
                    Else
                        fh = rs!elfechahasta
                    End If
                    
                    Dias = DateDiff("d", fd, fh) + 1
    
                    If Dias > 0 Then
                        sumaRanking = sumaRanking + (Dias * rs!valor)
                    End If
    
                    
                    rs.MoveNext
                Loop
                'Para crear la cabecera debo tener al menos un dato en el detalle
                If Not crearCabecera And sumaRanking > 0 Then
                    crearCabecera = True
                    StrSql = " INSERT INTO rep_lic_rank (licrandesc,bpronro,licranfecha,licranhora,tenro1,estrnro1,tenro2,estrnro2,tenro3,estrnro3) VALUES " & _
                             " ('" & historicoDesc & "'," & NroProceso & "," & ConvFecha(Date) & ",'" & Left(Time, 8) & "'," & _
                              tenro1 & "," & estrnro1 & "," & tenro2 & "," & estrnro2 & "," & tenro3 & "," & estrnro3 & ")"
                    objConn.Execute StrSql, , adExecuteNoRecords
                    licrannro = getLastIdentity(objConn, "rep_lic_rank")
                    Flog.writeline "Creada la cabecera para el reporte: " & NroProceso
                End If
                
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
                    If rsEstruc.EOF Then
                        sumaRanking = 0
                    End If
                End If
                    
                If sumaRanking > 0 Then
                    'inserto el detalle del empleado si hay puntaje
                    StrSql = " INSERT INTO rep_lic_rank_det (licrannro,ternro,empleg,nombre,puntaje,tenro1,estrnro1,tenro2,estrnro2,tenro3,estrnro3) VALUES " & _
                             " (" & licrannro & "," & rsEmp!Ternro & "," & rsEmp!empleg & ",'" & nombre & "'," & sumaRanking
                If tenro1 <> 0 Then
                    StrSql = StrSql & " ," & rsEstruc!tenro1
                    StrSql = StrSql & " ," & rsEstruc!estrnro1
                Else
                    StrSql = StrSql & " ,0,0"
                End If
                
                If tenro2 <> 0 Then
                    StrSql = StrSql & " ," & rsEstruc!tenro2
                    StrSql = StrSql & " ," & rsEstruc!estrnro2
                Else
                    StrSql = StrSql & " ,0,0"
                End If
                
                If tenro3 <> 0 Then
                    StrSql = StrSql & " ," & rsEstruc!tenro3
                    StrSql = StrSql & " ," & rsEstruc!estrnro3
                Else
                    StrSql = StrSql & " ,0,0"
                End If
                StrSql = StrSql & ")"
                objConn.Execute StrSql, , adExecuteNoRecords
                
                    Flog.writeline "Insertado el detalle para el empleado: " & rsEmp!Ternro & ", cabecera de reporte: " & licrannro & " y proceso: " & NroProceso
                Else
                    Flog.writeline "No hay puntaje para el empleado: " & rsEmp!Ternro
                End If
            Else
                Flog.writeline "No existen licencias a procesar para el empleado: " & rsEmp!Ternro & " en el rango de fechas " & fecDesde & " - " & fecHasta
            End If
            

            'Actualizo el progreso
            TiempoAcumulado = GetTickCount
            'CEmpleadosAProc = CEmpleadosAProc - 1
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



Sub analizarLicencia_old(ByRef rs As Recordset, ByVal fecDesde, ByVal fecHasta, ByRef sumaRanking As Double)
Dim mesDesdeLic
Dim mesHastaLic
Dim Dias
Dim fechaCiclo
Dim mesCiclo


    mesDesdeLic = Month(rs!elfechadesde)
    mesHastaLic = Month(rs!elfechahasta)
    Dias = 0
    'Caso donde la licencia cae en mismo mes
    If mesDesdeLic = mesHastaLic Then
        Dias = rs!elcantdias
        If rs!Dias < Dias Then
            sumaRanking = sumaRanking + ((Dias - rs!Dias) * rs!valor)
        End If
    Else
        'la licencia abarca mas de un mes analizo dia x dia
        fechaCiclo = rs!elfechadesde
        mesCiclo = Month(rs!elfechadesde)
        
        Do While fechaCiclo <= rs!elfechahasta
            If mesCiclo = Month(fechaCiclo) Then
                Dias = Dias + 1
            Else
                If rs!Dias < Dias Then
                    sumaRanking = sumaRanking + ((Dias - rs!Dias) * rs!valor)
                End If
                mesCiclo = Month(fechaCiclo)
                Dias = 1
            End If
            fechaCiclo = DateAdd("d", 1, fechaCiclo)
        Loop
        sumaRanking = sumaRanking + ((Dias - rs!Dias) * rs!valor)
    End If
        
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












