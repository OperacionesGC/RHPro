Attribute VB_Name = "MdlRHproRepControlHsExtras"
Option Explicit

Global Const Version = "1.00" 'Sebastian Stremel
Global Const FechaModificacion = "18/07/2014"
'Global Const UltimaModificacion = " CAS-24521 - PLA - Nuevo Reporte de Hs Extras "


Global NroProc As Long
Global objConnRep As New ADODB.Connection
Global HuboErrores As Boolean
Global teAreas As Integer
Global teUnidadNegocio As Integer

Private Sub Main()

Dim strCmdLine As String
Dim PID As String
Dim arrParametros
Dim Archivo As String
Dim nroReporte As Integer
Dim parametros As String
Dim rs_Batch_Proceso As New ADODB.Recordset


Dim legdesde As Long
Dim leghasta As Long
Dim estado As Integer
Dim Mes As String
Dim anio As Integer
Dim orden As String
Dim ArrParam
Dim titulofiltro As String

On Error GoTo ce
    strCmdLine = Command()
    arrParametros = Split(strCmdLine, " ", -1)
    If UBound(arrParametros) > 1 Then
        If IsNumeric(arrParametros(0)) Then
            NroProc = arrParametros(0)
            Etiqueta = arrParametros(1)
            EncriptStrconexion = CBool(arrParametros(2))
            c_seed = arrParametros(2)
        Else
            Exit Sub
        End If
    Else
        If UBound(arrParametros) > 0 Then
            If IsNumeric(arrParametros(0)) Then
                NroProc = arrParametros(0)
                Etiqueta = arrParametros(1)
            Else
                Exit Sub
            End If
        Else
            If IsNumeric(strCmdLine) Then
                NroProc = strCmdLine
            Else
                Exit Sub
            End If
        End If
    End If
    
    
    'Carga las configuraciones basicas, formato de fecha, string de conexion, tipo de BD y ubicacion del archivo de log
    Call CargarConfiguracionesBasicas
        
    'Creo el archivo de texto del desglose
    Archivo = PathFLog & "RHPro_RepControlHsExtras" & "-" & NroProc & ".log"
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Archivo, True)
        
    'Obtengo el Process ID
    PID = GetCurrentProcessId
    Flog.writeline "-------------------------------------------------"
    Flog.writeline "Version                  : " & Version
    Flog.writeline "Fecha Ultima Modificacion: " & FechaModificacion
    Flog.writeline "PID                      : " & PID
    Flog.writeline "-------------------------------------------------"
    Flog.writeline
    
    On Error Resume Next
    OpenConnection strconexion, objConn
    If Err.Number <> 0 Or Error_Encrypt Then
        Flog.writeline "Problemas con la conexión "
        Exit Sub
    End If
     
    OpenConnection strconexion, objConnProgreso
    If Err.Number <> 0 Or Error_Encrypt Then
        Flog.writeline "Problemas con la conexión "
        Exit Sub
    End If
    
    
    'Activo el manejador de errores
    On Error GoTo ce
        
    'Cambio el estado del proceso a Procesando
    StrSql = "UPDATE batch_proceso SET bprcpid = " & PID & ", bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcestado = 'Procesando', bprcprogreso = 0 WHERE bpronro = " & NroProc
    objConnProgreso.Execute StrSql, , adExecuteNoRecords
    
    Flog.writeline "Levanta Proceso y Setea Parámetros:  " & " " & Now
       
    'Obtiene los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE bpronro =" & NroProc
    OpenRecordset StrSql, rs_Batch_Proceso
           
    'Si no encuentra el proceso aborta el proceso
    If rs_Batch_Proceso.EOF Then
        Flog.writeline "No se encontro el registro en bach proceso. Nro de Proceso: " & NroProc
        Exit Sub
    End If
        
    'Obtiene la lsita de parametros
    parametros = rs_Batch_Proceso!bprcparam
    Flog.writeline "Parametros " & parametros
    
    
    ArrParam = Split(parametros, "@")
    legdesde = ArrParam(0)
    leghasta = ArrParam(1)
    estado = ArrParam(2)
    Mes = ArrParam(3)
    anio = ArrParam(4)
    orden = ArrParam(5)
    
    titulofiltro = " Bpronro: " & NroProc & " Leg Desde: " & legdesde & " Leg Hasta: " & leghasta & " MES: " & Mes & " Anio: " & anio
    'Inserto la cabecera
    OpenConnection strconexion, objConnRep
    StrSql = "INSERT INTO rep_controlHsExtras_cab "
    StrSql = StrSql & " (bpronro, titulofiltro, anio, mes) "
    StrSql = StrSql & " VALUES "
    StrSql = StrSql & " ( "
    StrSql = StrSql & NroProc
    StrSql = StrSql & ",'" & titulofiltro & "'"
    StrSql = StrSql & "," & anio
    StrSql = StrSql & "," & Mes
    StrSql = StrSql & " ) "
    objConnRep.Execute StrSql, , adExecuteNoRecords
    'llamar al reporte
    Call reporteControlHsExtras(legdesde, leghasta, estado, Mes, anio, orden)
    
    If Not HuboErrores Then
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcprogreso = 100, bprcestado = 'Procesado' WHERE bpronro = " & NroProc
    Else
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Incompleto' WHERE bpronro = " & NroProc
    End If
    objConnProgreso.Execute StrSql, , adExecuteNoRecords
        
    MyCommitTrans
    GoTo Fin
' ----------------------------------------------------------

Fin:
    StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcprogreso = 100, bprcestado = 'Procesado' WHERE bpronro = " & NroProc
    objConnProgreso.Execute StrSql, , adExecuteNoRecords
    Exit Sub
        
ce:
    MyRollbackTrans
    HuboErrores = True
    Flog.writeline " ------------------------------------------------------------"
    Flog.writeline Err.Description
    Flog.writeline "SQL: " & StrSql
    Flog.writeline " ------------------------------------------------------------"
    StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Incompleto' WHERE bpronro = " & NroProc
    objConnProgreso.Execute StrSql, , adExecuteNoRecords
End Sub

Private Sub reporteControlHsExtras(legdesde, leghasta, estado, Mes, anio, orden)

Dim StrSql As String
Dim rs_consultas As New ADODB.Recordset
Dim rs_consultas2 As New ADODB.Recordset
Dim listaAreas As String
Dim listaUnidadNegocio As String
Dim areas() As String
Dim negocios() As String
Dim area As Integer
Dim negocio As Integer
Dim cantDiasMes As Integer
Dim fechadesde As String
Dim fechahasta As String
Dim cantdias2 As Integer
Dim Ternro As Long
Dim Fecha As String
Dim horas_ausencia As Long
Dim arrDias()
Dim ternom As String
Dim ternom2 As String
Dim terape As String
Dim terape2 As String
Dim empleg As String
Dim estrUnidadNegocio As Integer
Dim estrArea As Integer
Dim k As Integer
Dim i As Integer
Dim j As Integer
Dim o As Integer
Dim l As Integer

Dim lngAlcanGrupo
Dim HsTeoricas As Long
Dim hs
Dim HsEnfermedad As Long
Dim HsReporte(11)
'variables confrep


Dim codHsAus As Integer
Dim codHsTeoricas As Integer
Dim codHsEnfermedad As String
Dim codhs(11)
Dim listaAreasInd As String

Dim progreso As Double
Dim porcentaje As Double
listaAreas = "0"
listaUnidadNegocio = "0"
On Error GoTo Desc
    'Levanto del confrep las estructuras
    StrSql = " SELECT * FROM confrep "
    StrSql = StrSql & " WHERE repnro=448 "
    OpenRecordset StrSql, rs_consultas
    If Not rs_consultas.EOF Then
        Do While Not rs_consultas.EOF
            Select Case rs_consultas!confnrocol
                Case 1:
                    teUnidadNegocio = IIf(EsNulo(rs_consultas!confval), 0, rs_consultas!confval)
                    listaUnidadNegocio = IIf(EsNulo(rs_consultas!confval2), "0", rs_consultas!confval2)
                Case 2:
                    teAreas = IIf(EsNulo(rs_consultas!confval), 0, rs_consultas!confval)
                    listaAreas = IIf(EsNulo(rs_consultas!confval2), "0", rs_consultas!confval2)
                Case 3:
                    codHsAus = IIf(EsNulo(rs_consultas!confval), 0, rs_consultas!confval)
                Case 14:
                    listaAreasInd = IIf(EsNulo(rs_consultas!confval2), "0", rs_consultas!confval2)
            End Select
            
        rs_consultas.MoveNext
        Loop
    Else
        Flog.writeline "No se configuro el confrep, se aborta el proceso. "
    End If
    
    
    'armo la fecha desde y hasta del mes
    fechadesde = "01/" & Right("00" & Mes, 2) & "/" & anio
    fechahasta = DateAdd("m", 1, fechadesde)
    fechahasta = DateAdd("d", -1, fechahasta)
    cantdias2 = DateDiff("d", fechadesde, fechahasta) + 1
    
    'PARA CADA UNA DE LAS AREAS Y LOS NEGOCIOS VOY ARMANDO LA TABLA
    areas = Split(listaAreas, ",")
    negocios = Split(listaUnidadNegocio, ",")
    progreso = 0
    If CDbl(UBound(areas)) > 1 Then
        porcentaje = 100 / CDbl(UBound(areas) + 1)
    Else
        porcentaje = 100 / 1
    End If
    For j = 0 To UBound(areas)
        area = areas(j)
        progreso = progreso + porcentaje
        For k = 0 To UBound(negocios)
            negocio = negocios(k)
            Fecha = fechadesde
            'Busco los empleados que cumplen con las estructuras
            StrSql = " SELECT empleado.ternro,empleado.empleg legajo, empleado.ternom + ' ' + empleado.ternom2 + ',' + empleado.terape + ' ' + empleado.terape2 nombre FROM empleado "
            StrSql = StrSql & " INNER JOIN his_estructura ON empleado.ternro = his_estructura.ternro AND his_estructura.tenro =" & teAreas & " and his_estructura.estrnro=" & area
            StrSql = StrSql & " INNER JOIN his_estructura est2 ON empleado.ternro = est2.ternro AND est2.tenro =" & teUnidadNegocio & " and est2.estrnro=" & negocio & " "
            StrSql = StrSql & " AND ( "
            StrSql = StrSql & " (his_estructura.htetdesde <= " & ConvFecha(fechadesde) & " AND (his_estructura.htethasta is null or his_estructura.htethasta >= " & ConvFecha(fechahasta)
            StrSql = StrSql & " or his_estructura.htethasta >= " & ConvFecha(fechadesde) & ")) OR "
            StrSql = StrSql & " (his_estructura.htetdesde >= " & ConvFecha(fechadesde) & " AND (his_estructura.htetdesde <= " & ConvFecha(fechahasta) & ")) "
            StrSql = StrSql & " )"
            StrSql = StrSql & " AND ("
            StrSql = StrSql & " (est2.htetdesde <= " & ConvFecha(fechadesde) & " AND (est2.htethasta is null or est2.htethasta >= " & ConvFecha(fechahasta)
            StrSql = StrSql & " or est2.htethasta >= " & ConvFecha(fechadesde) & ")) OR "
            StrSql = StrSql & " (est2.htetdesde >= " & ConvFecha(fechadesde) & " AND (est2.htetdesde <= " & ConvFecha(fechahasta) & "))"
            StrSql = StrSql & " )"
            StrSql = StrSql & " AND empleado.empleg >=" & legdesde & " AND empleado.empleg <=" & leghasta
            StrSql = StrSql & " AND empleado.empest =" & estado
            Flog.writeline "Query de empleado:" & StrSql
            OpenRecordset StrSql, rs_consultas
            If Not rs_consultas.EOF Then
                Do While Not rs_consultas.EOF
                    Ternro = rs_consultas!Ternro
                    Fecha = fechadesde
                    HsTeoricas = 0
                    HsEnfermedad = 0
                    'Busco los datos del empleado
                    StrSql = " SELECT empleg, ternom, ternom2, terape, terape2 FROM empleado "
                    StrSql = StrSql & " WHERE ternro=" & Ternro
                    OpenRecordset StrSql, rs_consultas2
                    If Not rs_consultas2.EOF Then
                        ternom = IIf(EsNulo(rs_consultas2!ternom), "", rs_consultas2!ternom)
                        ternom2 = IIf(EsNulo(rs_consultas2!ternom2), "", rs_consultas2!ternom2)
                        terape = IIf(EsNulo(rs_consultas2!terape), "", rs_consultas2!terape)
                        terape2 = IIf(EsNulo(rs_consultas2!terape2), "", rs_consultas2!terape2)
                        empleg = rs_consultas2!empleg
                    Else
                        Flog.writeline " No se encontraron los datos del empleado."
                        Exit Sub
                    End If
                    'hasta aca
                    'para cada dia del mes
                    For i = 1 To cantdias2
                        ReDim Preserve arrDias(i)
                        If tieneEstructura(Ternro, Fecha, area, negocio) Then
                            horas_ausencia = buscarHsAusencia(Ternro, Fecha, codHsAus)
                            arrDias(i) = arrDias(i) + CDbl(horas_ausencia)
                        Else
                            horas_ausencia = 0
                            arrDias(i) = arrDias(i) + CDbl(horas_ausencia)
                        End If
                        Fecha = DateAdd("d", 1, Fecha)
                    Next
                    
                    'inserto para cada dia un registro
                    StrSql = "INSERT INTO rep_controlHsExtras_det "
                    StrSql = StrSql & " ("
                    StrSql = StrSql & " bpronro "
                    StrSql = StrSql & ", Ternro  "
                    StrSql = StrSql & ", empleg  "
                    StrSql = StrSql & ", empape  "
                    StrSql = StrSql & ", empape2  "
                    StrSql = StrSql & ", empnom  "
                    StrSql = StrSql & ", empnom2  "
                    StrSql = StrSql & ", teUnidadNegocio  "
                    StrSql = StrSql & ", estrUnidadNegocio  "
                    StrSql = StrSql & ", teArea  "
                    StrSql = StrSql & ", estrArea  "
                    For i = 1 To cantdias2
                        StrSql = StrSql & ",dia" & i
                    Next
                    StrSql = StrSql & " ) "
                    StrSql = StrSql & " VALUES ( "
                    StrSql = StrSql & NroProc
                    StrSql = StrSql & "," & Ternro
                    StrSql = StrSql & "," & empleg
                    StrSql = StrSql & ",'" & terape & "'"
                    StrSql = StrSql & ",'" & terape2 & "'"
                    StrSql = StrSql & ",'" & ternom & "'"
                    StrSql = StrSql & ",'" & ternom2 & "'"
                    StrSql = StrSql & "," & teUnidadNegocio
                    StrSql = StrSql & "," & negocio
                    StrSql = StrSql & "," & teAreas
                    StrSql = StrSql & "," & area
                    For i = 1 To cantdias2
                        StrSql = StrSql & "," & arrDias(i)
                    Next
                    StrSql = StrSql & ")"
                    objConnRep.Execute StrSql, , adExecuteNoRecords
                    'hasta aca
                rs_consultas.MoveNext
                Loop
            Else
                'inserto para cada dia un registro
                StrSql = "INSERT INTO rep_controlHsExtras_det "
                StrSql = StrSql & " ("
                StrSql = StrSql & " bpronro "
                StrSql = StrSql & ", Ternro  "
                StrSql = StrSql & ", empleg  "
                StrSql = StrSql & ", teUnidadNegocio  "
                StrSql = StrSql & ", estrUnidadNegocio  "
                StrSql = StrSql & ", teArea  "
                StrSql = StrSql & ", estrArea  "
                For i = 1 To cantdias2
                    StrSql = StrSql & ",dia" & i
                Next
                StrSql = StrSql & " ) "
                StrSql = StrSql & " VALUES ( "
                StrSql = StrSql & NroProc
                StrSql = StrSql & ",0"
                StrSql = StrSql & ",0"
                StrSql = StrSql & "," & teUnidadNegocio
                StrSql = StrSql & "," & negocio
                StrSql = StrSql & "," & teAreas
                StrSql = StrSql & "," & area
                For i = 1 To cantdias2
                    StrSql = StrSql & ",0"
                Next
                StrSql = StrSql & ")"
                objConnRep.Execute StrSql, , adExecuteNoRecords
            End If
            rs_consultas.Close
        Next
        'UPDATE DEL PROGRESO
        StrSql = "UPDATE batch_proceso SET bprcprogreso =" & progreso & " WHERE bpronro = " & NroProc
        objConnProgreso.Execute StrSql, , adExecuteNoRecords
        'HASTA ACA

    Next
    
    'para cada una de las areas individuales busco los datos
    negocio = -1
    areas = Split(listaAreasInd, ",")
    For j = 0 To UBound(areas)
        area = areas(j)
        Fecha = fechadesde
        'Busco los empleados que cumplen con las estructuras
        StrSql = " SELECT empleado.ternro,empleado.empleg legajo, empleado.ternom + ' ' + empleado.ternom2 + ',' + empleado.terape + ' ' + empleado.terape2 nombre FROM empleado "
        StrSql = StrSql & " INNER JOIN his_estructura ON empleado.ternro = his_estructura.ternro AND his_estructura.tenro =" & teAreas & " and his_estructura.estrnro=" & area
        StrSql = StrSql & " AND ( "
        StrSql = StrSql & " (his_estructura.htetdesde <= " & ConvFecha(fechadesde) & " AND (his_estructura.htethasta is null or his_estructura.htethasta >= " & ConvFecha(fechahasta)
        StrSql = StrSql & " or his_estructura.htethasta >= " & ConvFecha(fechadesde) & ")) OR "
        StrSql = StrSql & " (his_estructura.htetdesde >= " & ConvFecha(fechadesde) & " AND (his_estructura.htetdesde <= " & ConvFecha(fechahasta) & ")) "
        StrSql = StrSql & " )"
        StrSql = StrSql & " AND empleado.empleg >=" & legdesde & " AND empleado.empleg <=" & leghasta
        StrSql = StrSql & " AND empleado.empest =" & estado
        OpenRecordset StrSql, rs_consultas
        If Not rs_consultas.EOF Then
            Do While Not rs_consultas.EOF
                Ternro = rs_consultas!Ternro
                Fecha = fechadesde
                HsTeoricas = 0
                HsEnfermedad = 0
                'Busco los datos del empleado
                StrSql = " SELECT empleg, ternom, ternom2, terape, terape2 FROM empleado "
                StrSql = StrSql & " WHERE ternro=" & Ternro
                OpenRecordset StrSql, rs_consultas2
                If Not rs_consultas2.EOF Then
                    ternom = IIf(EsNulo(rs_consultas2!ternom), "", rs_consultas2!ternom)
                    ternom2 = IIf(EsNulo(rs_consultas2!ternom2), "", rs_consultas2!ternom2)
                    terape = IIf(EsNulo(rs_consultas2!terape), "", rs_consultas2!terape)
                    terape2 = IIf(EsNulo(rs_consultas2!terape2), "", rs_consultas2!terape2)
                    empleg = rs_consultas2!empleg
                Else
                    Flog.writeline " No se encontraron los datos del empleado."
                    Exit Sub
                End If
                'hasta aca
                'para cada dia del mes
                For i = 1 To cantdias2
                    ReDim Preserve arrDias(i)
                    If tieneEstructura(Ternro, Fecha, area, negocio) Then
                        horas_ausencia = buscarHsAusencia(Ternro, Fecha, codHsAus)
                        arrDias(i) = arrDias(i) + CDbl(horas_ausencia)
                    Else
                        horas_ausencia = 0
                        arrDias(i) = arrDias(i) + CDbl(horas_ausencia)
                    End If
                    Fecha = DateAdd("d", 1, Fecha)
                Next
                
                'inserto para cada dia un registro
                StrSql = "INSERT INTO rep_controlHsExtras_det "
                StrSql = StrSql & " ("
                StrSql = StrSql & " bpronro "
                StrSql = StrSql & ", Ternro  "
                StrSql = StrSql & ", empleg  "
                StrSql = StrSql & ", empape  "
                StrSql = StrSql & ", empape2  "
                StrSql = StrSql & ", empnom  "
                StrSql = StrSql & ", empnom2  "
                StrSql = StrSql & ", teUnidadNegocio  "
                StrSql = StrSql & ", estrUnidadNegocio  "
                StrSql = StrSql & ", teArea  "
                StrSql = StrSql & ", estrArea  "
                For i = 1 To cantdias2
                    StrSql = StrSql & ",dia" & i
                Next
                StrSql = StrSql & " ) "
                StrSql = StrSql & " VALUES ( "
                StrSql = StrSql & NroProc
                StrSql = StrSql & "," & Ternro
                StrSql = StrSql & "," & empleg
                StrSql = StrSql & ",'" & terape & "'"
                StrSql = StrSql & ",'" & terape2 & "'"
                StrSql = StrSql & ",'" & ternom & "'"
                StrSql = StrSql & ",'" & ternom2 & "'"
                StrSql = StrSql & ",-1"
                StrSql = StrSql & ",-1"
                StrSql = StrSql & "," & teAreas
                StrSql = StrSql & "," & area
                For i = 1 To cantdias2
                    StrSql = StrSql & "," & arrDias(i)
                Next
                StrSql = StrSql & ")"
                objConnRep.Execute StrSql, , adExecuteNoRecords
                'hasta aca
            rs_consultas.MoveNext
            Loop
        Else
            'inserto los datos vacios
                StrSql = "INSERT INTO rep_controlHsExtras_det "
                StrSql = StrSql & " ("
                StrSql = StrSql & " bpronro "
                StrSql = StrSql & ", Ternro  "
                StrSql = StrSql & ", empleg  "
                StrSql = StrSql & ", teUnidadNegocio  "
                StrSql = StrSql & ", estrUnidadNegocio  "
                StrSql = StrSql & ", teArea  "
                StrSql = StrSql & ", estrArea  "
                For i = 1 To cantdias2
                    StrSql = StrSql & ",dia" & i
                Next
                StrSql = StrSql & " ) "
                StrSql = StrSql & " VALUES ( "
                StrSql = StrSql & NroProc
                StrSql = StrSql & ",0"
                StrSql = StrSql & ",0"
                StrSql = StrSql & ",-1"
                StrSql = StrSql & ",-1"
                StrSql = StrSql & "," & teAreas
                StrSql = StrSql & "," & area
                For i = 1 To cantdias2
                    StrSql = StrSql & ",0"
                Next
                StrSql = StrSql & ")"
                objConnRep.Execute StrSql, , adExecuteNoRecords
                'Hasta aca
                ReDim arrDias(0)
        End If
        rs_consultas.Close
        
    Next

    'hasta aca
    Exit Sub
    
Desc:
    Flog.Write Err.Description
    
End Sub

Function tieneEstructura(tercero, Dia, area, negocio)
Dim StrSql As String
Dim rsEstructura As New ADODB.Recordset

On Error GoTo ce
    'me fijo si el empleado tiene la estructura el dia
    StrSql = " SELECT empleado.ternro,empleado.empleg legajo, empleado.ternom + ' ' + empleado.ternom2 + ',' + empleado.terape + ' ' + empleado.terape2 nombre FROM empleado "
    StrSql = StrSql & " INNER JOIN his_estructura ON empleado.ternro = his_estructura.ternro AND his_estructura.tenro = " & teAreas & " and his_estructura.estrnro IN (" & area & ") "
    If negocio <> -1 Then
        StrSql = StrSql & " INNER JOIN his_estructura est2 ON empleado.ternro = est2.ternro AND est2.tenro =" & teUnidadNegocio & " and est2.estrnro=" & negocio & " "
    End If
    StrSql = StrSql & " AND  "
    StrSql = StrSql & " (his_estructura.htetdesde <= " & ConvFecha(Dia) & " AND (his_estructura.htethasta is null or his_estructura.htethasta >= " & ConvFecha(Dia)
    StrSql = StrSql & " ))"
    If negocio <> -1 Then
        StrSql = StrSql & " AND "
        StrSql = StrSql & " (est2.htetdesde <= " & ConvFecha(Dia) & " AND (est2.htethasta is null or est2.htethasta >= " & ConvFecha(Dia)
        StrSql = StrSql & " ))"
    End If
    StrSql = StrSql & " AND his_estructura.ternro=" & tercero
    OpenRecordset StrSql, rsEstructura
    If Not rsEstructura.EOF Then
        tieneEstructura = True
    Else
        tieneEstructura = False
    End If
    rsEstructura.Close
    Exit Function
ce:
    Flog.Write Err.Description
End Function

Function buscarHsAusencia(tercero, Fecha, tipoHora)
Dim StrSql As String
Dim rs_hs As New ADODB.Recordset
On Error GoTo ce
    StrSql = " SELECT * FROM gti_acumdiario "
    StrSql = StrSql & " WHERE ternro =" & tercero & " AND thnro=" & tipoHora
    StrSql = StrSql & " AND adfecha=" & ConvFecha(Fecha)
    OpenRecordset StrSql, rs_hs
    If Not rs_hs.EOF Then
        buscarHsAusencia = rs_hs!adcanthoras
    Else
        buscarHsAusencia = 0
    End If
    rs_hs.Close
    Exit Function

ce:
Flog.Write Err.Description
End Function


Function calcularHsTeoricas(tercero, fechadesde, fechahasta, codHsTeoricas)

Dim StrSql As String
Dim rs_hs As New ADODB.Recordset

On Error GoTo ce
'busca las hs teoricas de un empleado para cda dia
    StrSql = " SELECT horasdia FROM his_estructura "
    StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro AND his_estructura.tenro = 21 "
    StrSql = StrSql & " INNER JOIN regimenHorario ON regimenHorario.estrnro = his_estructura.estrnro "
    StrSql = StrSql & " WHERE his_estructura.tenro =" & codHsTeoricas & "  AND( "
    StrSql = StrSql & " (his_estructura.htetdesde <= " & ConvFecha(fechadesde) & " AND (his_estructura.htethasta is null or his_estructura.htethasta >= " & ConvFecha(fechahasta)
    StrSql = StrSql & " or his_estructura.htethasta >= " & ConvFecha(fechadesde) & ")) OR "
    StrSql = StrSql & " (his_estructura.htetdesde >= " & ConvFecha(fechadesde) & " AND (his_estructura.htetdesde <= " & ConvFecha(fechahasta) & ")) "
    StrSql = StrSql & " )"
    StrSql = StrSql & " AND his_estructura.ternro = " & tercero
    StrSql = StrSql & " ORDER BY his_estructura.htetdesde DESC "
    OpenRecordset StrSql, rs_hs
    If Not rs_hs.EOF Then
        calcularHsTeoricas = rs_hs!horasdia
    Else
        calcularHsTeoricas = 0
    End If
    rs_hs.Close
    
    Exit Function
ce:
Flog.Write Err.Description
End Function
