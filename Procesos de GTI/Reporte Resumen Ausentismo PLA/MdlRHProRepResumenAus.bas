Attribute VB_Name = "MdlRHProRepResumenAus"
Option Explicit

'Global Const Version = "1.00" 'Sebastian Stremel
'Global Const FechaModificacion = "02/07/2014"
'Global Const UltimaModificacion = " CAS-24597 - PLA - Reporte de Ausentismo "

Global Const Version = "1.1" 'Sebastian Stremel
Global Const FechaModificacion = "15/07/2014"
'Global Const UltimaModificacion = " CAS-24597 - PLA - Reporte de Ausentismo [Entrega 2]"
'Se buscan mas de un tipo de hora de ausencia

Global NroProc As Long
Global objConnRep As New ADODB.Connection
Global HuboErrores As Boolean


Private Sub Main()

Dim strCmdLine As String
Dim PID As String
Dim arrParametros
Dim Archivo As String
Dim nroReporte As Integer
Dim parametros As String
Dim rs_Batch_Proceso As New ADODB.Recordset
Dim rs As New ADODB.Recordset

Dim legdesde As Long
Dim leghasta As Long
Dim estado As Integer
Dim empresa As Long
Dim fechaDesde As String
Dim fechaHasta As String
Dim ArrParam
Dim titulofiltro As String
Dim nombreEmpresa As String


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
    Archivo = PathFLog & "RHPro_RepResumenAusentismo" & "-" & NroProc & ".log"
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
    empresa = ArrParam(3)
    fechaDesde = ArrParam(4)
    fechaHasta = ArrParam(5)
    
    'Busco el nombre de la empresa
    StrSql = " SELECT * FROM empresa "
    StrSql = StrSql & " WHERE estrnro=" & empresa
    OpenRecordset StrSql, rs
    If Not rs.EOF Then
        nombreEmpresa = rs!empnom
        Flog.writeline "La empresa es: " & nombreEmpresa
    Else
        nombreEmpresa = ""
        Flog.writeline "No se encontro el nombre de la empresa."
    End If
    rs.Close
    'hasta aca
    
    titulofiltro = " Bpronro: " & NroProc & " Leg Desde: " & legdesde & " A: " & leghasta & " Empresa: " & nombreEmpresa & " Fecha Desde: " & fechaDesde & " A: " & fechaHasta
    'Inserto la cabecera
    OpenConnection strconexion, objConnRep
    StrSql = "INSERT INTO rep_resumen_ausentismo_cab "
    StrSql = StrSql & " (bpronro, titulofiltro, fechaDesde, fechaHasta) "
    StrSql = StrSql & " VALUES "
    StrSql = StrSql & " ( "
    StrSql = StrSql & NroProc
    StrSql = StrSql & ",'" & titulofiltro & "'"
    StrSql = StrSql & ",'" & fechaDesde & "'"
    StrSql = StrSql & ",'" & fechaHasta & "'"
    StrSql = StrSql & " ) "
    objConnRep.Execute StrSql, , adExecuteNoRecords
    
    'llamar al reporte
    Call RepResumenAusentismo(legdesde, leghasta, estado, empresa, fechaDesde, fechaHasta)
    
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

Sub RepResumenAusentismo(ByVal legdesde As Long, ByVal leghasta As Long, ByVal estado As Integer, ByVal empresa As Long, ByVal fechaDesde As String, ByVal fechaHasta As String)
Dim rs_consultas As New ADODB.Recordset
Dim rs_licencias As New ADODB.Recordset
Dim rs_patologias As New ADODB.Recordset
Dim empleg As Long
Dim orden As Long
Dim Ternro As Long
Dim terape As String
Dim terape2 As String
Dim ternom As String
Dim ternom2 As String
Dim teArea As Long
Dim teSubArea As Long
Dim area As Long
Dim subarea As Long

Dim licDesde
Dim licHasta
Dim cantdias As Integer
Dim Dia As Date
Dim j As Integer
Dim listaLic As String

Dim progreso As Double
Dim porcentaje As Double
Dim StrSqlAux As String

Dim descarea As String
Dim descsubarea As String

Dim cantHs As Double
Dim descPatologias As String
Dim tipoHsAusencia As String

'Levanto los datos del confrep
StrSql = " SELECT * FROM confrep "
StrSql = StrSql & " WHERE repnro=442"
OpenRecordset StrSql, rs_consultas
If Not rs_consultas.EOF Then
    Do While Not rs_consultas.EOF
        Select Case rs_consultas!confnrocol
            Case 1:
                teArea = IIf(EsNulo(rs_consultas!confval), 0, rs_consultas!confval)
                Flog.writeline " Tipo Estructura Area: " & teArea
            Case 2:
                teSubArea = IIf(EsNulo(rs_consultas!confval), 0, rs_consultas!confval)
                Flog.writeline " Tipo Estructura SubArea: " & teSubArea
            Case 3:
                listaLic = IIf(EsNulo(rs_consultas!confval2), "", rs_consultas!confval2)
                Flog.writeline "Lista de licencias a buscar:" & rs_consultas!confval2
            Case 4:
                tipoHsAusencia = IIf(EsNulo(rs_consultas!confval2), "0", rs_consultas!confval2)
                Flog.writeline "Tipo de Hs configurado:" & rs_consultas!confval2
        End Select
    rs_consultas.MoveNext
    Loop
Else
    Flog.writeline " No esta configurado el confrep del reporte 442. "
End If
rs_consultas.Close

orden = 0
'Busco los empleados
StrSql = " SELECT empleado.ternro,empleado.empleg legajo, empleado.ternom, empleado.ternom2, empleado.terape, empleado.terape2, "
StrSql = StrSql & " area.estrnro area, subarea.estrnro subarea, estrarea.estrdabr descarea, estrsubarea.estrdabr descsubarea "
StrSql = StrSql & " FROM empleado "
If empresa <> -1 Then
    StrSql = StrSql & " INNER JOIN his_estructura empresa ON empleado.ternro = empresa.ternro AND empresa.tenro =10 And empresa.estrnro = " & empresa
End If
'busco el area del empleado
StrSql = StrSql & " LEFT JOIN his_estructura area ON empleado.ternro = area.ternro AND area.tenro =" & teArea
StrSql = StrSql & " LEFT JOIN estructura estrarea ON area.estrnro = estrarea.estrnro AND estrarea.tenro =" & teArea

'busco el subarea del empleado
StrSql = StrSql & " LEFT JOIN his_estructura subarea ON empleado.ternro = subarea.ternro AND subarea.tenro =" & teSubArea
StrSql = StrSql & " LEFT JOIN estructura estrsubarea ON subarea.estrnro = estrsubarea.estrnro AND estrsubarea.tenro =" & teSubArea
StrSql = StrSql & " WHERE "
If empresa <> -1 Then
    StrSql = StrSql & "("
    StrSql = StrSql & " (empresa.htetdesde <=  " & ConvFecha(fechaDesde) & " AND (empresa.htethasta is null or empresa.htethasta >= " & ConvFecha(fechaHasta)
    StrSql = StrSql & "or empresa.htethasta >=  " & ConvFecha(fechaDesde) & ")) OR"
    StrSql = StrSql & "(empresa.htetdesde >=  " & ConvFecha(fechaDesde) & " AND (empresa.htetdesde <= " & ConvFecha(fechaHasta) & "))"
    StrSql = StrSql & ") AND"
End If

StrSql = StrSql & " ((area.htetdesde <=  " & ConvFecha(fechaDesde) & " AND (area.htethasta is null or area.htethasta >= " & ConvFecha(fechaHasta)
StrSql = StrSql & "or area.htethasta >=  " & ConvFecha(fechaDesde) & ")) OR"
StrSql = StrSql & "(area.htetdesde >=  " & ConvFecha(fechaDesde) & " AND (area.htetdesde <= " & ConvFecha(fechaHasta) & "))"
StrSql = StrSql & ")"

StrSql = StrSql & " AND "
StrSql = StrSql & " ((subarea.htetdesde <=  " & ConvFecha(fechaDesde) & " AND (subarea.htethasta is null or subarea.htethasta >= " & ConvFecha(fechaHasta)
StrSql = StrSql & "or subarea.htethasta >=  " & ConvFecha(fechaDesde) & ")) OR"
StrSql = StrSql & "(subarea.htetdesde >=  " & ConvFecha(fechaDesde) & " AND (subarea.htetdesde <= " & ConvFecha(fechaHasta) & "))"
StrSql = StrSql & ")"
StrSql = StrSql & " AND empleado.empleg >=" & legdesde & " AND empleado.empleg <=" & leghasta
StrSql = StrSql & " AND empleado.empest =" & estado
OpenRecordset StrSql, rs_consultas
If Not rs_consultas.EOF Then
    porcentaje = 100 / rs_consultas.RecordCount
    Flog.writeline "Se encontraron empleados "
    Do While Not rs_consultas.EOF
        'para cada uno de los empleados busco los datos necesarios
        progreso = progreso + CDbl(porcentaje)
        orden = orden + 1
        empleg = IIf(EsNulo(rs_consultas!Legajo), 0, rs_consultas!Legajo) 'legajo
        Ternro = IIf(EsNulo(rs_consultas!Ternro), 0, rs_consultas!Ternro) 'nro de tercero
        terape = IIf(EsNulo(rs_consultas!terape), "", rs_consultas!terape) 'apellido
        terape2 = IIf(EsNulo(rs_consultas!terape2), "", rs_consultas!terape2) 'segundo apellido
        ternom = IIf(EsNulo(rs_consultas!ternom), "", rs_consultas!ternom) 'nombre
        ternom2 = IIf(EsNulo(rs_consultas!ternom2), "", rs_consultas!ternom2) 'segundo nombre
        area = IIf(EsNulo(rs_consultas!area), 0, rs_consultas!area) 'area del empleado
        subarea = IIf(EsNulo(rs_consultas!subarea), 0, rs_consultas!subarea) 'subarea del empleado
        descarea = IIf(EsNulo(rs_consultas!descarea), "", rs_consultas!descarea) 'area del empleado
        descsubarea = IIf(EsNulo(rs_consultas!descsubarea), "", rs_consultas!descsubarea) 'subarea del empleado
        
        'busco todas las licencias configuradas del empleado
        StrSqlAux = " SELECT elfechadesde, elfechahasta, elhoradesde, elhorahasta, elmaxhoras, empleado, tdnro, emp_lic.emp_licnro, eldiacompleto, elcantdias, eltipo, licestnro "
        StrSqlAux = StrSqlAux & " FROM emp_lic "
        StrSqlAux = StrSqlAux & " WHERE emp_lic.empleado=" & Ternro
        StrSqlAux = StrSqlAux & " AND tdnro IN (" & listaLic & ")"
        StrSqlAux = StrSqlAux & " AND ((elfechadesde <=  " & ConvFecha(fechaDesde) & " AND (elfechahasta is null or elfechahasta >= " & ConvFecha(fechaHasta)
        StrSqlAux = StrSqlAux & " or elfechahasta >=  " & ConvFecha(fechaDesde) & ")) OR"
        StrSqlAux = StrSqlAux & "(elfechadesde >=  " & ConvFecha(fechaDesde) & " AND (elfechadesde <= " & ConvFecha(fechaHasta) & ")))"
        OpenRecordset StrSqlAux, rs_licencias
        If Not rs_licencias.EOF Then
            Do While Not rs_licencias.EOF
                descPatologias = ""
                'busco las visitas y las patologias de las licencias
                StrSqlAux = " SELECT patologiadesabr FROM Licencia_visita "
                StrSqlAux = StrSqlAux & " INNER JOIN sopatol_visitas ON sopatol_visitas.visitamed = Licencia_visita.visitamed "
                StrSqlAux = StrSqlAux & " INNER JOIN sopatologias ON sopatologias.patologianro = sopatol_visitas.patologianro "
                StrSqlAux = StrSqlAux & " WHERE licencia_visita.emp_licnro=" & rs_licencias!emp_licnro
                OpenRecordset StrSqlAux, rs_patologias
                Do While Not rs_patologias.EOF
                    descPatologias = descPatologias & rs_patologias!patologiadesabr & " - "
                rs_patologias.MoveNext
                Loop
                rs_patologias.Close
                
                cantdias = DateDiff("d", fechaDesde, fechaHasta)
                StrSql = "INSERT INTO rep_resumen_ausentismo_det "
                StrSql = StrSql & "( bpronro, orden, ternro, empleg, empape, empape2, empnom, empnom2"
                StrSql = StrSql & ",area, subarea, DescArea, DescSubArea"
                Dia = fechaDesde
                For j = 0 To cantdias
                    StrSql = StrSql & ",dia" & Day(Dia)
                    Dia = DateAdd("d", 1, Dia)
                Next
                StrSql = StrSql & ", causaAusentismo, finicio, fAlta)"
                StrSql = StrSql & "VALUES"
                StrSql = StrSql & "(" & NroProc & ", "
                StrSql = StrSql & orden & ", "
                StrSql = StrSql & rs_consultas!Ternro & ", "
                StrSql = StrSql & rs_consultas!Legajo & ", "
                StrSql = StrSql & "'" & terape & "', "
                StrSql = StrSql & "'" & terape2 & "', "
                StrSql = StrSql & "'" & ternom & "', "
                StrSql = StrSql & "'" & ternom2 & "', "
                StrSql = StrSql & area & ", "
                StrSql = StrSql & subarea & ", "
                StrSql = StrSql & "'" & descarea & "', "
                StrSql = StrSql & "'" & descsubarea & "' "
                
                licDesde = rs_licencias!elfechadesde
                licHasta = rs_licencias!elfechahasta
                Dia = fechaDesde
                'si la licencia es dia completa, busco el valor del ausentismo del campo hs dias de la estructura gti
                Select Case rs_licencias!eltipo
                    Case 1: 'es dia completo
                        For j = 0 To cantdias
                            If CDate(Dia) >= CDate(licDesde) And CDate(Dia) <= CDate(licHasta) Then
                                StrSql = StrSql & ", " & buscarHsAusencia(Ternro, Dia, tipoHsAusencia)
                            Else
                                StrSql = StrSql & ",0 "
                            End If
                            Dia = DateAdd("d", 1, Dia)
                        Next
                    Case 2: 'Parcial Fija
                        For j = 0 To cantdias
                            If CDate(Dia) >= CDate(licDesde) And CDate(Dia) <= CDate(licHasta) Then
                                cantHs = CDbl(rs_licencias!elhorahasta) - CDbl(rs_licencias!elhoradesde)
                                If (CLng(Len(Trim(cantHs))) = CLng(4)) Then
                                    cantHs = Left(cantHs, 2) & "." & Right(cantHs, 2)
                                Else
                                    If Len(Trim(cantHs)) = 3 Then
                                        cantHs = Left(cantHs, 1) & "." & Right(cantHs, 2)
                                    End If
                                End If
                                StrSql = StrSql & ", " & cantHs
                            Else
                                StrSql = StrSql & ",0 "
                            End If
                            Dia = DateAdd("d", 1, Dia)
                        Next
                    Case 3: 'Parcial Variable
                        For j = 0 To cantdias
                            cantHs = CDbl(rs_licencias!elmaxhoras)
                            If CDate(Dia) >= CDate(licDesde) And CDate(Dia) <= CDate(licHasta) Then
                                StrSql = StrSql & ", " & cantHs
                            Else
                                StrSql = StrSql & ",0 "
                            End If
                            Dia = DateAdd("d", 1, Dia)
                        Next
                    
                End Select
                'busco la causa de ausentismo de la licencia (patologia)
                StrSql = StrSql & ",'" & descPatologias & "'"
                StrSql = StrSql & ",'" & rs_licencias!elfechadesde & "'"
                StrSql = StrSql & ",'" & rs_licencias!elfechahasta & "'"
                StrSql = StrSql & ")"
                objConnRep.Execute StrSql, , adExecuteNoRecords
            rs_licencias.MoveNext
            Loop
        Else
            Flog.writeline "No hay licencias para el empleado"
        End If
        'hasta aca
    rs_consultas.MoveNext
    'UPDATE DEL PROGRESO
    StrSql = "UPDATE batch_proceso SET bprcprogreso =" & progreso & " WHERE bpronro = " & NroProc
    objConnProgreso.Execute StrSql, , adExecuteNoRecords
    'HASTA ACA
    
    Loop
Else
    Flog.writeline "No hay empleados que cumplan con las condiciones: " & StrSql
End If
rs_consultas.Close
End Sub


Function buscarHsAusencia(Tercero, Fecha, tipoHora)
Dim StrSql As String
Dim rs_hs As New ADODB.Recordset
On Error GoTo ce
    StrSql = " SELECT * FROM gti_acumdiario "
    StrSql = StrSql & " WHERE ternro =" & Tercero & " AND thnro IN (" & tipoHora & ")"
    StrSql = StrSql & " AND adfecha=" & ConvFecha(Fecha)
    OpenRecordset StrSql, rs_hs
    Flog.writeline "Query Hs de ausencia:" & StrSql
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
