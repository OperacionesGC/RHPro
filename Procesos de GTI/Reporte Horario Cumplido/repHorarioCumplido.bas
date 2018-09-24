Attribute VB_Name = "repHorarioCumplido"
Option Explicit

'Global Const Version = "1.00"
'Global Const FechaModificacion = "21/12/2011"
'Global Const UltimaModificacion = "Inicial" 'sebastian stremel

'Global Const Version = "1.01"
'Global Const FechaModificacion = "10/05/2012"
'Global Const UltimaModificacion = "Se agrego la informacion si la registracion es manual y se corrgieron datos" 'Deluchi Ezequiel

'Global Const Version = "1.02"
'Global Const FechaModificacion = "22/05/2012"
'Global Const UltimaModificacion = "Se agrego totales horas convertidas y se corrigio consulta para filtro de agencias" '14769 - Deluchi Ezequiel

'Global Const Version = "1.03"
'Global Const FechaModificacion = "31/05/2012"
'Global Const UltimaModificacion = "Se corrigio cuando se aplicaba funcion cdbl() chequeos por vacio o nulo." '14769 - Deluchi Ezequiel

'Global Const Version = "1.04"
'Global Const FechaModificacion = "23/05/2013"
'Global Const UltimaModificacion = "Al generar el horario cumplido, ahora se toma en cuenta la fecha desde y hasta de gti_horcumplido" '19175 - Brzozowski Juan Pablo
 
'Global Const Version = "1.05"
'Global Const FechaModificacion = "17/06/2013"
'Global Const UltimaModificacion = "Se reemplazo la coma por punto cuando se hacia el update en la tabla rep_horariocumplidodet " 'CAS-20037 - Borrelli Facundo

'Global Const Version = "1.06"
'Global Const FechaModificacion = "04/07/2013"
'Global Const UltimaModificacion = "Se modifico la consulta para el filtro de horario cumplido donde se indican la fecha desde y hasta " 'CAS-20037 - Borrelli Facundo


'Global Const Version = "1.07"
'Global Const FechaModificacion = "18/07/2013"
'Global Const UltimaModificacion = "Se modifico la consulta que trae las registraciones de los empleados" 'CAS-20037 - AGD - ERROR EN FILTRO DE HORARIO CUMPLIDO [Entrega 2]- Fernandez, Matias- 18/07/2013

'Global Const Version = "1.08"
'Global Const FechaModificacion = "29/07/2013"
'Global Const UltimaModificacion = "Se modifico la consulta que trae las registraciones de los empleados" 'CAS-20037 - AGD - ERROR EN FILTRO DE HORARIO CUMPLIDO [Entrega 3]- Fernandez, Matias- 29/07/2013

'Global Const Version = "1.09"
'Global Const FechaModificacion = "20/08/2013"
'Global Const UltimaModificacion = "Comentario para nivelar fechas, solo se recompilo en rhdesa4" 'CAS-20037 - AGD - ERROR EN FILTRO DE HORARIO CUMPLIDO [Entrega 4]- Fernandez, Matias- 20/08/2013

'Global Const Version = "1.10"
'Global Const FechaModificacion = "04/10/2013"
'Global Const UltimaModificacion = "Se modifico la consulta al filtrar por agencia" 'CAS-21184 - AGD - ERROR EN FILTRO DE HORARIO CUMPLIDO- Borrelli, Facundo- 04/10/2013

Global Const Version = "1.11"
Global Const FechaModificacion = "05/06/2015"
Global Const UltimaModificacion = "Se cambiaron tipo de datos enteros a double"
                                  'Fernandez, Matias - CAS-31310 - AGD - Error en reporte de horario cumplido- se cambio el numero
                                  'de reporte de Integer a double
'---------------------------------------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------------------------------
Dim fs, f

Dim NroProceso As Long

Global Path As String
Global NArchivo As String
Global HuboErrores As Boolean
Global EmpErrores As Boolean

'Global Tabulador As Long
Global TiempoInicialProceso
Global TiempoAcumulado

Global IdUser As String
Global Fecha As Date
Global Hora As String
Global firmas As String
Global HoraConvertida As Double 'Se usa para guardar el valor de las horas transformadas a horas normales

Dim titulo As String
Dim legdesde As Long
Dim leghasta As Long
Dim fdesde As Date
Dim fhasta As Date
Dim agencia As Integer
Dim reloj As Integer
Dim tcontrato As String
Dim contrato As String
Dim l_salida As String
Dim cont As Integer
'Dim nrorep As Integer MDF 05/06/2015
Dim nrorep As Double
Dim tcontratoEmpleado As String
Dim nrocontratoEmpleado As String
Dim l_autorizadas
Dim l_rechazadas
Dim l_auto
Dim l_recha
Dim l_tipAutorizacion
Dim l_FichadasArr(20, 5)
Dim l_i
Dim l_empresa






'Dim filtroFirmas As String

Private Sub Main()

Dim NombreArchivo As String
Dim Directorio As String
Dim CArchivos
Dim archivo
Dim Folder
Dim strCmdLine As String
Dim Nombre_Arch As String

Dim StrSql As String
Dim objRs As New ADODB.Recordset
Dim cantRegistros
Dim PID As String
Dim parametros As String
Dim ArrParametros
'Dim progreso


'Dim empl_desde As Long
'Dim empl_hasta As Long
'Dim empl_estado As Integer
'Dim empresa As Long
'Dim tenro1 As Long
'Dim estrnro1 As Long
'Dim tenro2 As Long
'Dim estrnro2 As Long
'Dim tenro3 As Long
'Dim estrnro3 As Long
'Dim fecdesde As Date
'Dim fechasta As Date

    'strCmdLine = Command()
    'ArrParametros = Split(strCmdLine, " ", -1)
    'If UBound(ArrParametros) > 0 Then
    '    If IsNumeric(ArrParametros(0)) Then
    '        NroProceso = ArrParametros(0)
    '        Etiqueta = ArrParametros(1)
    '    Else
    '        Exit Sub
    '    End If
    'Else
    '    If IsNumeric(strCmdLine) Then
    '        NroProceso = strCmdLine
    '    Else
    '        Exit Sub
    '    End If
    'End If
        
    strCmdLine = Command()
    ArrParametros = Split(strCmdLine, " ", -1)
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
            If IsNumeric(strCmdLine) Then
                NroProceso = strCmdLine
            Else
                Exit Sub
            End If
        End If
    End If
    
    
    
    ' carga las configuraciones basicas, formato de fecha, string de conexion,
    ' tipo de BD y ubicacion del archivo de log
    Call CargarConfiguracionesBasicas

    TiempoInicialProceso = GetTickCount
    OpenConnection strconexion, objConn
    OpenConnection strconexion, objConnProgreso
    
    HuboErrores = False
    
    Nombre_Arch = PathFLog & "ReporteHorarioCumplidoProc" & "-" & NroProceso & ".log"
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Nombre_Arch, True)
    
    ' Obtengo el Process ID
    PID = GetCurrentProcessId
    Flog.writeline "-----------------------------------------------------------------"
    Flog.writeline "Version = " & Version
    Flog.writeline "Modificacion = " & UltimaModificacion
    Flog.writeline "Fecha = " & FechaModificacion
    Flog.writeline "-----------------------------------------------------------------"
    
    Flog.writeline
    Flog.writeline "PID = " & PID
    
    Flog.writeline "Inicio Proceso Control de Mano de Obra: " & Now
    Flog.writeline "Cambio el estado del proceso a Procesando"
    
    'Cambio el estado del proceso a Procesando
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcestado = 'Procesando', bprcpid = " & PID & ", bprcprogreso = 0 WHERE bpronro = " & NroProceso
    objConnProgreso.Execute StrSql, , adExecuteNoRecords
    
    Flog.writeline "Obtengo los datos del proceso"
    
    TiempoAcumulado = GetTickCount
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE bpronro = " & NroProceso
    OpenRecordset StrSql, objRs
    
    If Not objRs.EOF Then
       IdUser = objRs!IdUser
       Fecha = objRs!bprcfecha
       Hora = objRs!bprchora
       
       'Obtengo los parametros del proceso
       parametros = objRs!bprcparam
       ArrParametros = Split(parametros, "@")
       
       '-------------------------------------------------
       'titulo reporte
       titulo = ArrParametros(0)
       
       legdesde = ArrParametros(1)
       
       leghasta = ArrParametros(2)
       
       fdesde = ArrParametros(3)
       
       fhasta = ArrParametros(4)
       
       agencia = ArrParametros(5)
       
       reloj = ArrParametros(6)
       
       tcontrato = ArrParametros(7)
       
       contrato = ArrParametros(8)
       
       firmas = ArrParametros(9)
       
       Call GenerarReporte(titulo, legdesde, leghasta, fdesde, fhasta, agencia, reloj, tcontrato, contrato, firmas)
       
       
       
       '-------------------------------------------------
       
       'parametros viejos
       
       'Empleado - Desde Legajo
       'empl_desde = ArrParametros(0)
       
       'Empleado - Hasta Legajo
       'empl_hasta = ArrParametros(1)
       
       'Empleado - Estado
       'empl_estado = ArrParametros(2)
       
       'Empresa
       'empresa = ArrParametros(3)
       
       'Primer nivel organizacional
       'tenro1 = ArrParametros(4)
       'estrnro1 = ArrParametros(5)
       
       'Segundo nivel organizacional
       'tenro2 = ArrParametros(6)
       'estrnro2 = ArrParametros(7)
       
       'Tercero nivel organizacional
       'tenro3 = ArrParametros(8)
       'estrnro3 = ArrParametros(9)
       
       'Fecha desde
       'fecdesde = ArrParametros(10)
       
       'Fecha hasta
       'fechasta = ArrParametros(11)
       
       ' Proceso que genera los datos
       'Call GenerarDatos(empl_desde, empl_hasta, empl_estado, empresa, tenro1, estrnro1, tenro2, estrnro2, tenro3, estrnro3, fecdesde, fechasta)
       
    Else
       Exit Sub
    End If
    
    If objRs.State = adStateOpen Then objRs.Close
    
    'Actualizo el estado del proceso
    If Not HuboErrores Then
       StrSql = "UPDATE batch_proceso SET  bprcprogreso = 100, bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado' WHERE bpronro = " & NroProceso
       Flog.writeline "Proceso Finalizado Correctamente"
    Else
       StrSql = "UPDATE batch_proceso SET  bprcprogreso = 100, bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Incompleto' WHERE bpronro = " & NroProceso
       Flog.writeline "Proceso Incompleto"
    End If
    
    objConnProgreso.Execute StrSql, , adExecuteNoRecords
    
    Flog.writeline "Fin :" & Now
    Flog.Close

    Exit Sub
    
ce:
    HuboErrores = True
    Flog.writeline " Error: " & Err.Description & Now
End Sub
Function filtroFirmas08052012()
    filtroFirmas = ""
    If firmas = "0" Then 'rechazada
        filtroFirmas = " AND gti_registracion.regnro IN (select cysfircodext FROM cysfirmas where cystipnro = 33 AND cysfirrecha = -1 AND gti_registracion.regnro = cysfircodext) "
    End If
    If firmas = "-1" Then 'autorizadas
        'filtroFirmas = " AND gti_registracion.regnro IN (select cysfircodext FROM cysfirmas where cystipnro = 33 AND cysfirfin = -1 AND cysfiryaaut = -1 AND gti_registracion.regnro = cysfircodext) "
        filtroFirmas = " AND ( "
        filtroFirmas = filtroFirmas & "(gti_registracion.regnro IN (select cysfircodext FROM cysfirmas WHERE cystipnro = 33 AND cysfirfin = -1 AND cysfiryaaut = -1 AND gti_registracion.regnro = cysfircodext))"
        filtroFirmas = filtroFirmas & " OR (gti_registracion.regnro NOT IN (select cysfircodext FROM cysfirmas WHERE cystipnro = 33 AND gti_registracion.regnro = cysfircodext)))"
    End If
    If firmas = "2" Then 'Pendientes (Sin autorizar)
        filtroFirmas = " AND gti_registracion.regnro NOT IN (select cysfircodext FROM cysfirmas where cystipnro = 33 AND ((cysfirfin = -1 AND cysfiryaaut = -1) OR (cysfirrecha = -1)) AND gti_registracion.regnro = cysfircodext) "
        filtroFirmas = filtroFirmas & " AND gti_registracion.regnro IN (select cysfircodext FROM cysfirmas where cystipnro = 33 AND ((cysfirfin = 0 AND cysfiryaaut = 0 AND cysfirrecha = 0)) AND gti_registracion.regnro = cysfircodext) "
    End If

End Function

Function filtroFirmas()
    filtroFirmas = ""
    If firmas = "0" Then 'rechazada
        'filtroFirmas = " AND gti_registracion.regnro IN (select cysfircodext FROM cysfirmas where cystipnro = 33 AND cysfirrecha = -1 AND gti_registracion.regnro = cysfircodext) "
        filtroFirmas = " AND (gti_horcumplido.regent IN (select cysfircodext FROM cysfirmas where cystipnro = 33 AND cysfirrecha = -1 AND gti_horcumplido.regent = cysfircodext) "
        filtroFirmas = filtroFirmas & "OR gti_horcumplido.regsal IN (select cysfircodext FROM cysfirmas where cystipnro = 33 AND cysfirrecha = -1 AND gti_horcumplido.regsal = cysfircodext)) "

    End If
    If firmas = "-1" Then 'autorizadas
        'filtroFirmas = " AND gti_registracion.regnro IN (select cysfircodext FROM cysfirmas where cystipnro = 33 AND cysfirfin = -1 AND cysfiryaaut = -1 AND gti_registracion.regnro = cysfircodext) "
        filtroFirmas = " AND ( "
        filtroFirmas = filtroFirmas & "(gti_horcumplido.regsal IN (select cysfircodext FROM cysfirmas WHERE cystipnro = 33 AND cysfirfin = -1 AND cysfiryaaut = -1 AND gti_horcumplido.regsal = cysfircodext))"
        filtroFirmas = filtroFirmas & "AND (gti_horcumplido.regent IN (select cysfircodext FROM cysfirmas WHERE cystipnro = 33 AND cysfirfin = -1 AND cysfiryaaut = -1 AND gti_horcumplido.regent = cysfircodext))"
        filtroFirmas = filtroFirmas & " OR (gti_horcumplido.regsal NOT IN (select cysfircodext FROM cysfirmas WHERE cystipnro = 33 AND gti_horcumplido.regsal = cysfircodext)))"
    End If
    If firmas = "2" Then 'Pendientes (Sin autorizar)
        filtroFirmas = " AND ( (gti_horcumplido.regsal NOT IN (select cysfircodext FROM cysfirmas where cystipnro = 33 AND ((cysfirfin = -1 AND cysfiryaaut = -1) OR (cysfirrecha = -1)) AND gti_horcumplido.regsal = cysfircodext) "
        filtroFirmas = filtroFirmas & " AND gti_horcumplido.regsal IN (select cysfircodext FROM cysfirmas where cystipnro = 33 AND ((cysfirfin = 0 AND cysfiryaaut = 0 AND cysfirrecha = 0)) AND gti_horcumplido.regsal = cysfircodext)) "
    
        filtroFirmas = filtroFirmas & " AND ( gti_horcumplido.regent NOT IN (select cysfircodext FROM cysfirmas where cystipnro = 33 AND ((cysfirfin = -1 AND cysfiryaaut = -1) OR (cysfirrecha = -1)) AND gti_horcumplido.regent = cysfircodext) "
        filtroFirmas = filtroFirmas & " AND gti_horcumplido.regent IN (select cysfircodext FROM cysfirmas where cystipnro = 33 AND ((cysfirfin = 0 AND cysfiryaaut = 0 AND cysfirrecha = 0)) AND gti_horcumplido.regent = cysfircodext)) )"
    
    End If

End Function

Function Horas(ByVal ths, ByVal Etiqueta, ByVal Ternro, ByVal Fecha, ByVal relnro, ByVal cont, ByVal regsal, ByVal EtiqAlfaNum)
    Dim l_salida
    Dim l_sql As String
    Dim StrSql As String
    Dim sql As String
    Dim rsconsult As New ADODB.Recordset
    Dim valor As Double
    'Dim HoraConvertida As Double 'Se usa para guardar el valor de las horas transformadas a horas normales
    'Dim cont As Integer
    
    
    l_sql = "SELECT SUM(horcant) cant "
    l_sql = l_sql & " FROM gti_horcumplido "
    l_sql = l_sql & " LEFT JOIN gti_reg_comp ON gti_horcumplido.regsal = gti_reg_comp.regnro "
    l_sql = l_sql & " INNER JOIN gti_registracion ON gti_registracion.regnro = gti_horcumplido.regsal "
    l_sql = l_sql & " WHERE gti_horcumplido.ternro = " & Ternro
    l_sql = l_sql & " AND horfecrep = " & Fecha
    l_sql = l_sql & " AND thnro IN ( " & ths & ")"
    l_sql = l_sql & " AND relnro = " & relnro
    l_sql = l_sql & " AND gti_registracion.regnro IN (" & regsal & ") "
    If tcontratoEmpleado <> "" Then
        l_sql = l_sql & " AND tipocont = '" & tcontratoEmpleado & "' "
        
    Else
        l_sql = l_sql & " AND ( tipocont = '' OR tipocont is null)"
    End If
    If nrocontratoEmpleado <> "" Then
        l_sql = l_sql & " AND cont = " & nrocontratoEmpleado
    Else
        l_sql = l_sql & " AND cont is null "
    End If
        
    'response.Write(l_sql)
    'cont = cont + 1
    OpenRecordset l_sql, rsconsult
    If cont <= 9 Then                   'la columna 10 se usa para total de horas convertidas
        If ths = 1 Then                 'tipo de hora 1 siempre es hora normal
            HoraConvertida = 0
        End If
        
        If IsNull(rsconsult!Cant) Then
            valor = 0
        Else
            valor = rsconsult!Cant
        End If
        
        'Transformamos las horas segun el porcentaje
        If Not IsNull(EtiqAlfaNum) Then
            If CStr(EtiqAlfaNum) <> "0" And CStr(LTrim(RTrim(EtiqAlfaNum))) <> "" Then
                HoraConvertida = HoraConvertida + valor * (1 + (CDbl(EtiqAlfaNum) / 100))
            Else
                If CStr(EtiqAlfaNum) = "0" Then
                    HoraConvertida = HoraConvertida + valor
                End If
            End If
        End If
        'FB- Se reemplazo en el campo que tenia la cantidad de horas, la coma por un punto.
        'Ejemplo: 4,67 al reemplazarlo se obtiene 4.67, se reemplazo porque en el update la coma se tomaba como que separaba dos campos. Replace(valor, ",", ".")
        StrSql = " UPDATE rep_horariocumplidodet set campo" & cont & " = '" & Etiqueta & "', valor" & cont & "=" & Replace(valor, ",", ".") & ","
        StrSql = StrSql & " campo10 = 'Horas Convertidas', valor10 = " & HoraConvertida
        StrSql = StrSql & " WHERE fecha= " & Fecha & " AND ternro =" & Ternro & " and repnro = " & nrorep & " AND empresa= " & "'" & l_empresa & "'" & " AND contrato= " & "'" & tcontratoEmpleado & " - " & nrocontratoEmpleado & "'"
        objConn.Execute StrSql, , adExecuteNoRecords
        Flog.writeline "se actualizo la hora & " & StrSql
    Else
        Flog.writeline "Se acabo el limite de tipos de horas"
    End If
    'Select Case ths
    '    Case 1:
    '        Horas = Replace(rsconsult!confval2, ",", ".")
    'End Select
    
    
    'If Not rsconsult.EOF Then
    '    If Not IsNull(l_rs2(0)) Then
    '        l_salida = FormatNumber(l_rs2(0), 2)
    '    Else
    '        l_salida = FormatNumber(0, 2)
    '    End If
    'Else
    '    l_salida = FormatNumber(0, 2)
    'End If
    'Horas = l_salida
    rsconsult.Close
End Function


Sub Firmadas()
    Dim l_i
    Dim l_sql3 As String
    Dim rsconsult3 As New ADODB.Recordset
    For l_i = 0 To UBound(l_FichadasArr)
    l_autorizadas = ""
    l_rechazadas = ""
    
        If ((l_FichadasArr(l_i, 0) <> "" Or l_FichadasArr(l_i, 1) <> "") And (l_FichadasArr(l_i, 1))) Then
            If l_FichadasArr(l_i, 2) <> "" Then
                l_sql3 = " SELECT cysfirfin, cysfirautoriza, cysfirdestino, cysfiryaaut, cysfirrecha FROM cysfirmas "
                l_sql3 = l_sql3 & " WHERE cysfirmas.cystipnro = " & l_tipAutorizacion & " AND cysfirmas.cysfircodext = " & l_FichadasArr(l_i, 2)
                l_sql3 = l_sql3 & " ORDER BY cysfirsecuencia DESC "
                OpenRecordset l_sql3, rsconsult3
                If Not rsconsult3.EOF Then
                    If rsconsult3!cysfiryaaut = 0 Then
                        l_auto = "No"
                    Else
                        l_auto = "Si"
                    End If
                    If rsconsult3!cysfirrecha = -1 Then
                        l_recha = "Si"
                    Else
                        l_recha = "No"
                    End If
                Else
                    l_auto = "Si" 'Cambio la logica, si no se cargo manualmente, esta autorizada.
                    l_recha = "No"
                End If
                rsconsult3.Close
                l_autorizadas = l_autorizadas & l_auto
                l_rechazadas = l_rechazadas & l_recha
                
            Else
                l_autorizadas = l_autorizadas & "Si"
                l_rechazadas = l_rechazadas & "No"
            End If
           
        End If
        l_FichadasArr(l_i, 3) = l_autorizadas
        l_FichadasArr(l_i, 4) = l_rechazadas

    Next
    
    
    'For l_i = 0 To UBound(l_fichadasarr, 1)
    '    If l_fichadasarr(l_i, 0) <> "" Or l_fichadasarr(l_i, 1) <> "" Then
    '        If l_fichadasarr(l_i, 2) <> "" Then
    '            l_sql = " SELECT cysfirfin, cysfirautoriza, cysfirdestino, cysfiryaaut, cysfirrecha FROM cysfirmas "
    '            l_sql = l_sql & " WHERE cysfirmas.cystipnro = " & l_tipAutorizacion & " AND cysfirmas.cysfircodext = " & l_fichadasarr(l_i, 2)
    '            l_sql = l_sql & " ORDER BY cysfirsecuencia DESC "
    '            OpenRecordset l_sql3, rsconsult3
                'response.Write(l_sql)
                'response.End()
    '            If Not rsconsult3.EOF Then
    '                If rsconsult3!cysfiryaaut = 0 Then
    '                    l_auto = "No"
    '                Else
    '                    l_auto = "Si"
    '                End If
    '                If rsconsult3!cysfirrecha = -1 Then
    '                    l_recha = "Si"
    '                Else
    '                    l_recha = "No"
    '                End If
    '            Else
    '                l_auto = "Si" 'Cambio la logica, si no se cargo manualmente, esta autorizada.
    '                l_recha = "No"
    '            End If
    '            rsconsult3.Close
    '            l_autorizadas = l_autorizadas & l_auto
    '            l_rechazadas = l_rechazadas & l_recha
    '        Else
    '            l_autorizadas = l_autorizadas & "Si"
    '            l_rechazadas = l_rechazadas & "No"
    '        End If
    '    End If
    'Next


End Sub

Function Fichadas(ByVal ths, ByVal Ternro, ByVal Fecha, ByVal relnro, regsal)
    
    Dim l_cont
    Dim l_autorizadas
    Dim l_rechazadas
    Dim l_sql As String
    Dim l_sql3 As String
    Dim StrSql As String
    Dim rsconsult As New ADODB.Recordset
    Dim rsconsult3 As New ADODB.Recordset
    Dim objRs As New ADODB.Recordset
    Dim i As Integer
    Dim l_sql4 As String
    Dim rsconsult4 As New ADODB.Recordset
    l_autorizadas = ""
    l_rechazadas = ""
    'ReDim l_FichadasArr(20, 2) '0=E - 1=S - 2=regnro(S)
    'Dim l_i
    l_i = 0
    
    Flog.writeline " Buscando las fichadas " & Fecha
    'l_sql = " SELECT gti_registracion.reghora, regentsal, fechaproc, regfecha, reghora, gti_registracion.regnro  "
    'l_sql = l_sql & " FROM gti_horcumplido "
    'l_sql = l_sql & " INNER JOIN gti_registracion ON gti_registracion.regnro = gti_horcumplido.regent "
    'l_sql = l_sql & " LEFT JOIN gti_reg_comp ON gti_horcumplido.regent = gti_reg_comp.regnro "
    'l_sql = l_sql & " WHERE gti_horcumplido.ternro = " & Ternro
    'l_sql = l_sql & " AND horfecrep = " & Fecha
    'l_sql = l_sql & " AND relnro = " & relnro
    
    
    l_sql = " SELECT DISTINCT gti_registracion.reghora, regentsal, fechaproc, regfecha, reghora, gti_registracion.regnro, gti_horcumplido.regsal, gti_registracion.regmanual "
    l_sql = l_sql & " From gti_horcumplido "
    l_sql = l_sql & " INNER JOIN gti_registracion ON gti_registracion.regnro = gti_horcumplido.regent or gti_registracion.regnro = gti_horcumplido.regsal "
    l_sql = l_sql & " LEFT JOIN gti_reg_comp ON gti_horcumplido.regent = gti_reg_comp.regnro "
    l_sql = l_sql & " WHERE gti_horcumplido.ternro = " & Ternro
    l_sql = l_sql & " AND horfecrep = " & Fecha
    l_sql = l_sql & " AND relnro = " & relnro
    
    If tcontratoEmpleado <> "" Then
        l_sql = l_sql & " AND tipocont = '" & tcontratoEmpleado & "' "
    Else
        l_sql = l_sql & " AND ( tipocont = '' OR tipocont is null)"
    End If
    If nrocontratoEmpleado <> "" Then
        l_sql = l_sql & " AND cont = " & nrocontratoEmpleado
    Else
        l_sql = l_sql & " AND cont is null "
    End If
    l_sql = l_sql & filtroFirmas
    'l_sql = l_sql & " UNION "
    'l_sql = l_sql & " SELECT gti_registracion.reghora, regentsal, fechaproc, regfecha, reghora, gti_registracion.regnro   "
    'l_sql = l_sql & " FROM gti_horcumplido "
    'l_sql = l_sql & " INNER JOIN gti_registracion ON gti_registracion.regnro = gti_horcumplido.regsal "
    'l_sql = l_sql & " LEFT JOIN gti_reg_comp ON gti_horcumplido.regsal = gti_reg_comp.regnro "
    'l_sql = l_sql & " WHERE gti_horcumplido.ternro = " & Ternro
    'l_sql = l_sql & " AND horfecrep = " & Fecha
    'l_sql = l_sql & " AND relnro = " & relnro
    If tcontratoEmpleado <> "" Then
        l_sql = l_sql & " AND tipocont = '" & tcontratoEmpleado & "' "
    Else
        l_sql = l_sql & " AND ( tipocont = '' OR tipocont is null)"
    End If
    If nrocontratoEmpleado <> "" Then
        l_sql = l_sql & " AND cont = " & nrocontratoEmpleado
    Else
        l_sql = l_sql & " AND cont is null "
    End If
    l_sql = l_sql & filtroFirmas
    l_sql = l_sql & " ORDER BY fechaproc, regfecha, gti_registracion.reghora "
    OpenRecordset l_sql, rsconsult
    If Not rsconsult.EOF Then
        l_salida = ""
        'For i = 0 To UBound(l_FichadasArr)
        '    l_FichadasArr(i, 0) = ""
        '    l_FichadasArr(i, 1) = ""
        '    l_FichadasArr(i, 2) = ""
        '    l_FichadasArr(i, 3) = ""
        '    l_FichadasArr(i, 4) = ""
        '    l_FichadasArr(i, 5) = ""
        'Next
        regsal = "0"
        
        Do While Not rsconsult.EOF
            
            'armo arreglo para fichadas
            
            'If rsconsult!regentsal = "E" Then
            '    If l_FichadasArr(l_i, 0) <> "" Then
            '        l_i = l_i + 1
            '    End If
            '    l_FichadasArr(l_i, 0) = rsconsult!reghora
                'response.Write(l_FichadasArr(l_i,0))
            'End If
            'If rsconsult!regentsal = "S" Then
            '    If l_FichadasArr(l_i, 1) <> "" Then
            '        l_i = l_i + 1
            '    End If
            '    l_FichadasArr(l_i, 1) = rsconsult!reghora
                'response.Write(l_FichadasArr(l_i,1))
            '    l_FichadasArr(l_i, 2) = rsconsult!Regnro
                'response.Write(l_FichadasArr(l_i,2))
                'l_FichadasArr(l_i, 5) = rsconsult!regentsal
            'End If
            'l_FichadasArr(l_i, 5) = rsconsult!regentsal
            
            'Firmadas
            'l_salida = l_salida & l_FichadasArr(l_i, 5) & "," & l_FichadasArr(l_i, 0) & "," & l_FichadasArr(l_i, 1) & "," & l_FichadasArr(l_i, 2) & "," & l_FichadasArr(l_i, 3) & "," & l_FichadasArr(l_i, 4) & "@"
            'hasta aca
            'l_salida = l_salida & rsconsult!regentsal & "," & rsconsult!reghora & "," & l_autorizadas & "," & l_rechazadas & "@"
            
            'me fijo si esta autorizada cada una
            If l_tipAutorizacion <> "" Then
                l_sql4 = " SELECT cysfirfin, cysfirautoriza, cysfirdestino, cysfiryaaut, cysfirrecha FROM cysfirmas "
                l_sql4 = l_sql4 & " WHERE cysfirmas.cystipnro = " & l_tipAutorizacion & " AND cysfirmas.cysfircodext = " & rsconsult!Regnro
                l_sql4 = l_sql4 & " ORDER BY cysfirsecuencia DESC "
                OpenRecordset l_sql4, rsconsult4
                l_auto = ""
                l_recha = ""
                If Not rsconsult4.EOF Then
                    If rsconsult!Regnro <> "" Then
                        If rsconsult4!cysfiryaaut = 0 Then
                            l_auto = "No"
                        Else
                            l_auto = "Si"
                        End If
                        If rsconsult4!cysfirrecha = -1 Then
                            l_recha = "Si"
                        Else
                            l_recha = "No"
                        End If
                    End If
                Else
                    l_auto = "Si" 'Cambio la logica, si no se cargo manualmente, esta autorizada.
                    l_recha = "No"
                End If
                rsconsult4.Close
           Else
                Flog.writeline "no esta configurado el circuito de autorizacion"
           End If
            l_salida = l_salida & rsconsult!regentsal & "," & rsconsult!reghora & "," & l_auto & "," & l_recha & "," & rsconsult!regmanual & "@"
            If rsconsult!regentsal = "S" Then
                regsal = regsal & "," & rsconsult!regsal
                'regmanual = rsconsult!regmanual
            End If
        rsconsult.MoveNext
        Loop
        'Firmadas
        'l_salida = ""
        'For l_i = 0 To UBound(l_FichadasArr)
        '    l_salida = l_salida & l_FichadasArr(l_i, 5) & "," & l_FichadasArr(l_i, 0) & "," & l_FichadasArr(l_i, 1) & "," & l_FichadasArr(l_i, 2) & "," & l_FichadasArr(l_i, 3) & "," & l_FichadasArr(l_i, 4) & "@"
        'Next
        
    Else
        l_salida = ""
        Flog.writeline "No tiene fichada"
    End If

    Fichadas = ""
    rsconsult.Close
End Function


Sub GenerarReporte(ByVal titulo As String, ByVal legdesde As Long, ByVal leghasta As Long, ByVal fdesde As Date, ByVal fhasta As Date, ByVal agencia As Integer, ByVal reloj As Integer, ByVal tcontrato As String, ByVal contrato As String, ByVal firmas As Integer)
Dim StrSql As String
Dim rsconsult As New ADODB.Recordset
Dim rsconsult1 As New ADODB.Recordset
Dim rsconsult2 As New ADODB.Recordset

Dim rs As New ADODB.Recordset
Dim usuario As String
Dim con_seg As Boolean
Dim l_sql As String



Dim CEmpleadosAProc As Long
Dim IncPorc As Double
Dim Progreso As Double

Dim sql As String
On Error GoTo MError

MyBeginTrans

'codigo
Dim l_regsal
Dim l_regmanual
Dim l_HayAutorizacion  'Es para ver si las autorizaciones estan activas
Dim l_PuedeVer         'Es para ver si las autorizaciones estan activas
Dim l_auto
Dim l_recha
Dim l_empleg
Dim l_nombre
Dim l_nombre2
Dim l_terape
Dim l_terape2

'Dim l_salida
 l_sql = "select cystipo.* from cystipo "
 l_sql = l_sql & "where (cystipo.cystipact = -1) and cystipo.cystipnro = 33 "
 
 OpenRecordset l_sql, rsconsult
 
 l_HayAutorizacion = Not rsconsult.EOF
 
 If Not rsconsult.EOF Then
    l_tipAutorizacion = 33
 End If
 rsconsult.Close
 
 
'-------------Firmas-----------

'cargo el confrep
'l_sql = "SELECT COUNT(DISTINCT confnrocol) "
'l_sql = l_sql & " FROM confrep "
'l_sql = l_sql & " WHERE repnro = 267 "
'l_sql = l_sql & " AND conftipo = 'TH' "
'OpenRecordset l_sql, rsConsult
'If Not rsConsult.EOF Then
'    l_cant_col = l_rs(0)
'    l_nro_col = l_nro_col + l_cant_col
'Else
'    l_cant_col = 0
'End If
'rsConsult.Close


'hasta aca




'inserto en la tabla encabezado
StrSql = " INSERT INTO rep_horariocumplido (bpronro,titulo,legdesde,leghasta,fdesde,fhasta,agencia,reloj,tipocontrato,firmas) "
StrSql = StrSql & " VALUES (" & NroProceso & "," & "'" & titulo & "'," & legdesde & ", " & leghasta & ", " & ConvFecha(fdesde) & ", "
StrSql = StrSql & "" & ConvFecha(fhasta) & "," & "'" & agencia & "'," & reloj & ", " & "'" & tcontrato & "', " & firmas & ") "
objConn.Execute StrSql, , adExecuteNoRecords

'reviso si tiene algun reloj configurado
StrSql = "SELECT count(*) cant FROM gti_rel_usu "
StrSql = StrSql & " WHERE UPPER(gti_rel_usu.iduser) = '" & UCase(IdUser) & "'"
OpenRecordset StrSql, rsconsult

If rsconsult.EOF Then
    con_seg = False
Else
    If rsconsult!Cant > 0 Then
        con_seg = True
    Else
        con_seg = False
    End If
End If
rsconsult.Close

'agrego codigo para optimizar consultas sql
Dim l_sql_from
Dim l_sql_where

l_sql_from = ""
l_sql_where = ""

'------------------------------------------------------------------------------
If firmas = "0" Then 'rechazada
   l_sql_from = " INNER JOIN cysfirmas ON gti_registracion.regnro = cysfircodext AND cystipnro = 33 AND cysfirrecha = -1 "
End If

If firmas = "-1" Then 'autorizadas
    l_sql_from = "LEFT JOIN cysfirmas ON gti_registracion.regnro = cysfircodext and  cystipnro = 33  "
  '  l_sql_where = " AND ((cystipnro = 33 AND cysfirfin = -1 AND cysfiryaaut = -1 ) or (cysfircodext is null))"
    l_sql_where = "AND ((cysfirfin = -1 and cysfiryaaut = -1 and cysfirrecha = 0) or (cysfirfin is null and cysfiryaaut is null and cysfirrecha is null) )"
End If

If firmas = "2" Then 'Pendientes (Sin autorizar)
    l_sql_from = " INNER JOIN cysfirmas ON gti_registracion.regnro = cysfircodext AND (cystipnro = 33 AND cysfirfin = 0 AND cysfiryaaut = 0 AND cysfirrecha = 0)  "
End If
'---------------------------------------------------------------------------

    
    StrSql = "SELECT distinct empleado.ternro, empleado.empleg, terape, ternom "
    StrSql = StrSql & " ,tipocont, cont, gti_registracion.relnro, horfecrep, relcodext, reldabr  "
    If (agencia <> "0") Then
        StrSql = StrSql & " , agencia.htethasta "
    End If
    StrSql = StrSql & " FROM empleado "
    StrSql = StrSql & " INNER JOIN gti_registracion ON gti_registracion.ternro = empleado.ternro "
    StrSql = StrSql & " LEFT JOIN gti_reg_comp ON gti_registracion.regnro = gti_reg_comp.regnro "
    StrSql = StrSql & " INNER JOIN gti_horcumplido ON gti_registracion.regnro = gti_horcumplido.regsal "
    'JPB - Se agrega contro en la fecha desde y hasta de la tabla gti_horcumplido
    'StrSql = StrSql & " AND  hordesde >=  " & ConvFecha(fdesde) & "  AND  horhasta  <= " & ConvFecha(fhasta)
    
    If con_seg Then 'solo de los relojes que el usuario puede ver -------
        StrSql = StrSql & " INNER JOIN gti_rel_usu ON gti_registracion.relnro = gti_rel_usu.relnro AND upper(gti_rel_usu.iduser) = '" & UCase(IdUser) & "'"
    End If
    
    StrSql = StrSql & " INNER JOIN gti_reloj ON gti_registracion.relnro = gti_reloj.relnro "
    
    'filtro from
    StrSql = StrSql & l_sql_from
    
    'Filtros por agencia
    StrSql = StrSql & " INNER JOIN his_estructura agencia ON empleado.ternro = agencia.ternro  AND agencia.tenro = 28 "
    'l_sql = l_sql & " AND (agencia.htetdesde<= " & cambiafecha(date(), "YMD", true) & " AND (agencia.htethasta is null or agencia.htethasta>= " & cambiafecha(date(), "YMD", true) & "))"
    'l_sql = l_sql & " AND (agencia.htetdesde <= " & cambiafecha(l_hasta, "YMD", true) & " AND (agencia.htethasta <= " & cambiafecha(l_hasta, "YMD", true) & " OR agencia.htethasta IS NULL)) "
    
    If (agencia <> "0") Then
        StrSql = StrSql & " AND ( "
        StrSql = StrSql & " (agencia.htetdesde <= " & ConvFecha(fdesde) & " AND (agencia.htethasta >= " & ConvFecha(fdesde) & " OR agencia.htethasta IS NULL)) "
        StrSql = StrSql & " OR "
        StrSql = StrSql & " (agencia.htetdesde >= " & ConvFecha(fdesde) & "  AND (agencia.htetdesde <= " & ConvFecha(fhasta) & " OR agencia.htetdesde IS NULL))"
        StrSql = StrSql & " ) "
        StrSql = StrSql & " AND agencia.estrnro = " & agencia
    End If
    
    StrSql = StrSql & " WHERE 1=1 "
    
    'Filtro Firmas - Condicion WHERE
    StrSql = StrSql & l_sql_where
    
    StrSql = StrSql & " AND empleado.empleg >= " & legdesde & " AND empleado.empleg <= " & leghasta
    
    If tcontrato <> "" Then
        StrSql = StrSql & " AND tipocont = '" & tcontrato & "' "
    End If
    
     'FB - 04/07/2013 - Se modificó el rango de fechas para utilizar en el reporte, se necesitaba que muestre el reporte
     'ingresando la misma fecha desde y hasta y se tuvieron en cuenta todas las posibilidades
    'StrSql = StrSql & "    AND (hordesde >=  " & ConvFecha(fdesde) & "  AND horhasta <=  " & ConvFecha(fhasta) & ")"
    'StrSql = StrSql & "    OR (horhasta  <= " & ConvFecha(fhasta) & " AND horhasta >= " & ConvFecha(fhasta) & ")"
    'StrSql = StrSql & "    OR (hordesde  >= " & ConvFecha(fdesde) & " AND hordesde <= " & ConvFecha(fdesde) & ")"
    'StrSql = StrSql & "    OR (horhasta  <= " & ConvFecha(fhasta) & " AND horhasta >= " & ConvFecha(fhasta) & "))"
    
    If contrato <> "" Then
        StrSql = StrSql & " AND cont = " & contrato
    End If
    
    If reloj <> 0 Then
        StrSql = StrSql & " AND gti_registracion.relnro = " & reloj
    End If
    
    If (agencia <> "0") Then
        StrSql = StrSql & " AND horfecgen >= ( case when " & ConvFecha(fdesde) & "  <  agencia.htetdesde THEN  agencia.htetdesde  ELSE " & ConvFecha(fdesde) & "  END )"
        'StrSql = StrSql & " AND horfecgen >= ( case when agencia.htethasta is null THEN  " & ConvFecha(fdesde) & " ELSE agencia.htetdesde END )"
        StrSql = StrSql & " AND  horfecgen  <= ( case when agencia.htethasta is null THEN  " & ConvFecha(fhasta) & " ELSE agencia.htethasta END ) "
    'FB - 04/10/2013 - Se quito el else
    'Else
    '    StrSql = StrSql & " AND horfecgen >= " & ConvFecha(fdesde)
    '    StrSql = StrSql & " AND  horfecgen  <= " & ConvFecha(fhasta)
    End If
    
    'FB - 04/10/2013 - Se agrego la consulta del else fuera del else
    StrSql = StrSql & " AND horfecgen >= " & ConvFecha(fdesde)
    StrSql = StrSql & " AND  horfecgen  <= " & ConvFecha(fhasta)
    'StrSql = StrSql & filtroFirmas
    'l_sql = l_sql & filtro()
    'contrato, empresa, legajo, apellido, fecha
    StrSql = StrSql & " ORDER BY tipocont, "
    StrSql = StrSql & " cont, gti_registracion.relnro, empleado.empleg, terape, horfecrep asc" ', horfecrep
    OpenRecordset StrSql, rsconsult
    
    'empresa
    'l_empresa = rsconsult!reldabr
    'tcontratoEmpleado = ""
    'nrocontratoEmpleado = ""
    'tcontratoEmpleado = rsconsult!tipocont
    'nrocontratoEmpleado = rsconsult!contrato
    'para cada empleado empiezo a registrar el detalle
    
    'datos del procesando

    'Determino la proporcion de progreso
     Progreso = 0
     CEmpleadosAProc = rsconsult.RecordCount
    If CEmpleadosAProc = 0 Then
       CEmpleadosAProc = 1
       'Flog.writeline "No hay empleados para procesar"
    End If
    IncPorc = (99 / CEmpleadosAProc)
    
    If rsconsult.EOF Then 'si no hay empleados salgo del proceso
        Flog.writeline "No hay empleados para procesar"
        Exit Sub
    End If
    
    Do While Not rsconsult.EOF
        'Incremento el progreso
            Progreso = Progreso + IncPorc
        'empresa
        l_empresa = rsconsult!reldabr
        tcontratoEmpleado = ""
        nrocontratoEmpleado = ""
        
        If Not EsNulo(rsconsult!tipocont) Then
            tcontratoEmpleado = rsconsult!tipocont
        End If
        
        'tcontratoEmpleado = rsconsult!tipocont
        If Not EsNulo(rsconsult!cont) Then
            nrocontratoEmpleado = rsconsult!cont
        End If
        
        Call Fichadas(1, rsconsult!Ternro, ConvFecha(rsconsult!horfecrep), rsconsult!relnro, l_regsal)
        'aca insert en tabla detalle
        sql = " SELECT * FROM rep_horariocumplido WHERE bpronro= " & NroProceso 'busco el nro de reporte
        OpenRecordset sql, rsconsult1
        If Not rsconsult1.EOF Then
            nrorep = rsconsult1!nrorep
        End If
        rsconsult1.Close
        
        'busco la empresa
        'sql = " select * from gti_reloj  where relnro=" & reloj
        'OpenRecordset sql, rsconsult2
        'If Not rsconsult2.EOF Then
            
        'Else
        '    l_empresa = ""
        '    Flog.writeline "NO SE ENCONTRO LA EMPRESA"
        'End If
        'rsconsult2.Close
        'hasta aca
        
        'busco los datos del empleado
        sql = "SELECT * FROM empleado WHERE ternro=" & rsconsult!Ternro
        OpenRecordset sql, rsconsult1
        If Not rsconsult1.EOF Then
            l_empleg = rsconsult1!empleg
            l_nombre = rsconsult1!ternom
            If IsNull(rsconsult1!ternom2) Then
                l_nombre2 = ""
                
            Else
                l_nombre2 = rsconsult1!ternom2
            End If
            l_terape = rsconsult1!terape
            
            If IsNull(rsconsult1!terape2) Then
                l_terape2 = ""
                
            Else
                l_terape2 = rsconsult1!terape2
            End If
        Else
            Flog.writeline "NO SE ENCONTRARON LOS DATOS DEL EMPLEADO "
        End If
        rsconsult1.Close
        'hasta aca
        
        'INSERTO LOS DATOS GENERALES EN LA TABLA DETALLE
        StrSql = "INSERT INTO rep_horariocumplidodet (repnro,ternro,contrato,empresa,empleado,ape1,ape2,nom1,nom2,fecha,fichadas)"
        StrSql = StrSql & " VALUES (" & nrorep & ", "
        StrSql = StrSql & "'" & rsconsult!Ternro & "',"
        StrSql = StrSql & "upper('" & rsconsult!tipocont & " - " & rsconsult!cont & "'),"
        StrSql = StrSql & "'" & l_empresa & "',"
        StrSql = StrSql & l_empleg & ", "
        StrSql = StrSql & "'" & l_nombre & "', "
        StrSql = StrSql & "'" & l_nombre2 & "', "
        StrSql = StrSql & "'" & l_terape & "' ,"
        StrSql = StrSql & "'" & l_terape2 & "', "
        StrSql = StrSql & ConvFecha(rsconsult!horfecrep) & ", "
        StrSql = StrSql & "'" & l_salida & "')"
        
        objConn.Execute StrSql, , adExecuteNoRecords
        
        'busco los tipos de hora en el confrep
        StrSql = " SELECT distinct confnrocol,* FROM confrep WHERE repnro = 267 and conftipo= 'TH' "
        OpenRecordset StrSql, rsconsult1
        cont = 0
        If rsconsult1.EOF Then
            Flog.writeline "No hay configurados tipo de hora en el confrep"
            GoTo MError
        End If
        
        Do While Not rsconsult1.EOF
            cont = cont + 1
            Call Horas(IIf(EsNulo(rsconsult1!confval), "0", rsconsult1!confval), IIf(EsNulo(rsconsult1!confetiq), "0", rsconsult1!confetiq), rsconsult!Ternro, ConvFecha(rsconsult!horfecrep), IIf(EsNulo(rsconsult!relnro), "0", rsconsult!relnro), IIf(EsNulo(cont), "0", cont), IIf(EsNulo(l_regsal), "0", l_regsal), IIf(EsNulo(rsconsult1!confval2), "0", rsconsult1!confval2))
            
            rsconsult1.MoveNext
        Loop
        rsconsult1.Close
        'contrato = rsconsult!cont
        'tcontrato = rsconsult!tipocont
            
            
            
    rsconsult.MoveNext
            
                'TiempoInicialProceso = GetTickCount
                'MyBeginTrans
                    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso
                    StrSql = StrSql & " , bprctiempo = " & TiempoInicialProceso
                    StrSql = StrSql & " WHERE bpronro = " & NroProceso
                    objConnProgreso.Execute StrSql, , adExecuteNoRecords
                'MyCommitTrans
                
        TiempoAcumulado = GetTickCount
        StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso
        StrSql = StrSql & ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'"
        StrSql = StrSql & " WHERE bpronro = " & NroProceso
        'StrSql = StrSql & ", bprcempleados ='" & CStr(cantidadProcesada) & "' WHERE bpronro = " & NroProceso
        objConnProgreso.Execute StrSql, , adExecuteNoRecords
    Loop
    
    GoTo seguir
    
    
MError:
    MyRollbackTrans
'    Resume Next
    Flog.writeline
    Flog.writeline "***************************************************************"
    Flog.writeline " Error: " & Err.Description
    Flog.writeline " Última Sql ejecutada: " & StrSql
    Flog.writeline "***************************************************************"
    Flog.writeline
    HuboErrores = True
    Exit Sub
seguir:
    'tcontrato = 1
    MyCommitTrans
    
End Sub
'--------------------------------------------------------------------
' Se encarga de generar los datos
'--------------------------------------------------------------------
Sub GenerarDatos(ByVal desde As Long, ByVal hasta As Long, ByVal estado As Integer, ByVal empresa As Long, ByVal tenro1 As Long, ByVal estrnro1 As Long, ByVal tenro2 As Long, ByVal estrnro2 As Long, ByVal tenro3 As Long, ByVal estrnro3 As Long, ByVal fecdesde As Date, ByVal fechasta As Date)

Dim StrSql As String
Dim rsconsult As New ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim objFeriado As New Feriado

'Variables donde se guardan los datos del INSERT final
Dim EmpCuit As String
Dim EmpNombre As String
Dim EmpDire As String
Dim EmpTernro As String
Dim EmpLogo As String
Dim EmpLogoAlto
Dim EmpLogoAncho

Dim CoefResta As Double
Dim CoefHsNormal As Double
Dim CoefHsPeAct As Double
Dim lista_HsExtras As String
Dim lista_ResPuesto As String
Dim lista_PagoHoras As String
Dim lista_NoPagoHoras_S As String
Dim lista_NoPagoHoras_R As String

Dim fecCalculo As Date

Dim tedabr1 As String
Dim estrdabr1 As String
Dim tedabr2 As String
Dim estrdabr2 As String
Dim tedabr3 As String
Dim estrdabr3 As String

Dim estrnro1_ant As Long
Dim estrnro2_ant As Long
Dim estrnro3_ant As Long
Dim estrdabr1_ant As String
Dim estrdabr2_ant As String
Dim estrdabr3_ant As String
Dim guardar As Boolean
Dim hsextras As Double
Dim resvpuesto As Double
Dim pgohs As Double
Dim nopgohs As Double

Dim hsNormal
Dim hspreact
Dim tothsaus
Dim pgohsporc
Dim nopgohsporc
Dim hspres
Dim hspresporc
Dim hsextrasporc
Dim toths
Dim tothsporc
                        
Dim cantemp As Integer
Dim dias_Hab As Integer
Dim contarDiasHab As Boolean
Dim esFeriado As Boolean
Dim insertarRegistro As Boolean

Dim Progreso As Double
Dim cantidadProcesada As Integer
Dim IncPorc As Double

Dim fecdesde_empl As Date
Dim fechasta_empl As Date
Dim orden As Integer
Dim tothsestr As Double
Dim total_AR As Double
Dim total_ANR As Double
Dim confetiq_ant As String


Dim Cargar_registro As Boolean
Dim encontre_rango As Boolean
Dim i As Integer

On Error GoTo MError

    MyBeginTrans
    
    '-------------------------------------------------------------------------
    ' Busco los datos de la empresa
    '--------------------------------------------------------------------------
    StrSql = "SELECT empresa.empnom,empresa.ternro, detdom.calle,nro,codigopostal, localidad.locdesc "
    StrSql = StrSql & " FROM empresa "
    StrSql = StrSql & " LEFT JOIN cabdom ON empresa.ternro = cabdom.ternro "
    StrSql = StrSql & " LEFT JOIN detdom ON detdom.domnro = cabdom.domnro "
    StrSql = StrSql & " LEFT JOIN localidad ON detdom.locnro = localidad.locnro "
    StrSql = StrSql & " WHERE empresa.estrnro = " & empresa
    OpenRecordset StrSql, rsconsult
    If rsconsult.EOF Then
        Flog.writeline Espacios(Tabulador * 1) & "***************************************************************"
        Flog.writeline Espacios(Tabulador * 1) & "Error. No se encontro la empresa."
        Flog.writeline Espacios(Tabulador * 1) & "***************************************************************"
        GoTo Fin_error
    Else
        EmpNombre = rsconsult!empnom
        EmpDire = rsconsult!calle & " " & rsconsult!nro & "<br>" & rsconsult!codigopostal & " " & rsconsult!locdesc
        EmpTernro = rsconsult!Ternro
    End If
    rsconsult.Close
    
    '-------------------------------------------------------------------------
    'Consulta para obtener el cuit de la empresa
    '-------------------------------------------------------------------------
    StrSql = "SELECT cuit.nrodoc FROM tercero " & _
             " INNER JOIN ter_doc cuit ON (tercero.ternro = cuit.ternro and cuit.tidnro = 6)" & _
             " Where tercero.ternro =" & EmpTernro
    OpenRecordset StrSql, rsconsult
    If rsconsult.EOF Then
        Flog.writeline Espacios(Tabulador * 1) & "No se encontró el CUIT de la Empresa."
        EmpCuit = "&nbsp;"
    Else
        EmpCuit = rsconsult!nrodoc
    End If
    rsconsult.Close
    
    '-------------------------------------------------------------------------
    'Consulta para buscar el logo de la empresa
    '-------------------------------------------------------------------------
    StrSql = "SELECT ter_imag.terimnombre, tipoimag.tipimdire, tipoimag.tipimanchodef, tipoimag.tipimaltodef" & _
        " FROM ter_imag " & _
        " INNER JOIN tipoimag ON tipoimag.tipimnro = 1 AND tipoimag.tipimnro = ter_imag.tipimnro" & _
        " AND ter_imag.ternro =" & EmpTernro
    OpenRecordset StrSql, rsconsult
    If rsconsult.EOF Then
        Flog.writeline Espacios(Tabulador * 1) & "No se encontró el Logo de la Empresa."
        EmpLogo = ""
        EmpLogoAlto = 0
        EmpLogoAncho = 0
    Else
        EmpLogo = rsconsult!tipimdire & rsconsult!terimnombre
        EmpLogoAlto = rsconsult!tipimaltodef
        EmpLogoAncho = rsconsult!tipimanchodef
    End If
    rsconsult.Close
    
    
    '-------------------------------------------------------------------------
    'Busco la descripcion del Primer Nivel Organizacional
    '-------------------------------------------------------------------------
    If tenro1 <> 0 Then
        StrSql = "SELECT tenro,tedabr FROM tipoestructura "
        StrSql = StrSql & " WHERE tipoestructura.tenro = " & tenro1
        OpenRecordset StrSql, rsconsult
        If Not rsconsult.EOF Then
            tedabr1 = rsconsult!tedabr
        End If
        rsconsult.Close
        
        If estrnro1 <> -1 Then
            StrSql = "SELECT estrnro,estrdabr FROM estructura "
            StrSql = StrSql & " WHERE estrnro = " & estrnro1
            OpenRecordset StrSql, rsconsult
            If Not rsconsult.EOF Then
                estrdabr1 = rsconsult!estrdabr
            End If
            rsconsult.Close
        End If
    End If
    
    '-------------------------------------------------------------------------
    'Busco la descripcion del Segundo Nivel Organizacional
    '-------------------------------------------------------------------------
    If tenro2 <> 0 Then
        StrSql = "SELECT tenro,tedabr FROM tipoestructura "
        StrSql = StrSql & " WHERE tipoestructura.tenro = " & tenro2
        OpenRecordset StrSql, rsconsult
        If Not rsconsult.EOF Then
            tedabr2 = rsconsult!tedabr
        End If
        rsconsult.Close
        
        If estrnro2 <> -1 Then
            StrSql = "SELECT estrnro,estrdabr FROM estructura "
            StrSql = StrSql & " WHERE estrnro = " & estrnro2
            OpenRecordset StrSql, rsconsult
            If Not rsconsult.EOF Then
                estrdabr2 = rsconsult!estrdabr
            End If
            rsconsult.Close
        End If
    End If
    
    '-------------------------------------------------------------------------
    'Busco la descripcion del tercer Nivel Organizacional
    '-------------------------------------------------------------------------
    If tenro3 <> 0 Then
        StrSql = "SELECT tenro,tedabr FROM tipoestructura "
        StrSql = StrSql & " WHERE tipoestructura.tenro = " & tenro3
        OpenRecordset StrSql, rsconsult
        If Not rsconsult.EOF Then
            tedabr3 = rsconsult!tedabr
        End If
        rsconsult.Close
        
        If estrnro3 <> -1 Then
            StrSql = "SELECT estrnro,estrdabr FROM estructura "
            StrSql = StrSql & " WHERE estrnro = " & estrnro3
            OpenRecordset StrSql, rsconsult
            If Not rsconsult.EOF Then
                estrdabr3 = rsconsult!estrdabr
            End If
            rsconsult.Close
        End If
    End If
    
    '-------------------------------------------------------------------------
    'Inserto la cabecera del reporte
    '-------------------------------------------------------------------------
    StrSql = "INSERT INTO rep_mano_obra  (bpronro,empldesde,emplhasta,emplest,empnombre,empdire,empcuit,emplogo,"
    StrSql = StrSql & "emplogoalto,emplogoancho,tenro1,tedabr1,estrnro1,estrdabr1,tenro2,tedabr2,estrnro2,estrdabr2,"
    StrSql = StrSql & "tenro3,tedabr3,estrnro3,estrdabr3,fecdesde,fechasta,fecha,hora,IdUser) VALUES ("
    StrSql = StrSql & NroProceso & ","
    StrSql = StrSql & desde & ","
    StrSql = StrSql & hasta & ","
    StrSql = StrSql & estado & ","
    StrSql = StrSql & "'" & EmpNombre & "',"
    StrSql = StrSql & "'" & EmpDire & "',"
    StrSql = StrSql & "'" & EmpCuit & "',"
    StrSql = StrSql & "'" & EmpLogo & "',"
    StrSql = StrSql & EmpLogoAlto & ","
    StrSql = StrSql & EmpLogoAncho & ","
    StrSql = StrSql & tenro1 & ","
    StrSql = StrSql & "'" & tedabr1 & "',"
    StrSql = StrSql & estrnro1 & ","
    StrSql = StrSql & "'" & estrdabr1 & "',"
    StrSql = StrSql & tenro2 & ","
    StrSql = StrSql & "'" & tedabr2 & "',"
    StrSql = StrSql & estrnro2 & ","
    StrSql = StrSql & "'" & estrdabr2 & "',"
    StrSql = StrSql & tenro3 & ","
    StrSql = StrSql & "'" & tedabr3 & "',"
    StrSql = StrSql & estrnro3 & ","
    StrSql = StrSql & "'" & estrdabr3 & "',"
    StrSql = StrSql & ConvFecha(fecdesde) & ","
    StrSql = StrSql & ConvFecha(fechasta) & ","
    StrSql = StrSql & ConvFecha(Fecha) & ","
    StrSql = StrSql & "'" & Hora & "',"
    StrSql = StrSql & "'" & IdUser & "')"
    
    objConn.Execute StrSql, , adExecuteNoRecords
    
    '-------------------------------------------------------------------------
    'Busco la configuración del reporte
    '-------------------------------------------------------------------------
    StrSql = "SELECT * FROM confrep WHERE repnro = 171"
    OpenRecordset StrSql, rsconsult
    
    Flog.writeline " "
    Flog.writeline "***************************************************************"
    Flog.writeline "Buscando valores en la configuración del reporte (confrep). Los tipos válidos son:"
    Flog.writeline "        CRE - Coeficiente Resta (V.AlfaNum)"
    Flog.writeline "        CHN - Coeficiente Horas Normal (V.AlfaNum)"
    Flog.writeline "        CHP - Coeficiente Horas Pe.Act (V.AlfaNum)"
    Flog.writeline "        TH - Tipos de Horas"
    Flog.writeline "             Columna 5  - Horas Reserva Puesto"
    Flog.writeline "             Columna 7  - Horas Pago"
    Flog.writeline "             columna 8  - Horas No Pago - Accion SUMA"
    Flog.writeline "             Columna 8  - Horas No Pago - Accion RESTA. Se le resta el CRE"
    Flog.writeline "             Columna 11 - Horas Extras"
    
    CoefResta = 0
    CoefHsNormal = 1
    CoefHsPeAct = 1
    lista_HsExtras = "0"
    lista_ResPuesto = "0"
    lista_PagoHoras = "0"
    lista_NoPagoHoras_S = "0"
    lista_NoPagoHoras_R = "0"
    
    Do Until rsconsult.EOF
        Select Case rsconsult!conftipo
            Case "CRE":
                CoefResta = Replace(rsconsult!confval2, ",", ".")
            Case "CHN":
                CoefHsNormal = Replace(rsconsult!confval2, ",", ".")
            Case "CHP":
                CoefHsPeAct = Replace(rsconsult!confval2, ",", ".")
            Case "TH":
                Select Case rsconsult!confnrocol
                    Case 11:
                        lista_HsExtras = lista_HsExtras & "," & rsconsult!confval
                    Case 5:
                        lista_ResPuesto = lista_ResPuesto & "," & rsconsult!confval
                    Case 7:
                        lista_PagoHoras = lista_PagoHoras & "," & rsconsult!confval
                    Case 8:
                        If rsconsult!confaccion = "sumar" Then
                            lista_NoPagoHoras_S = lista_NoPagoHoras_S & "," & rsconsult!confval
                        Else
                            lista_NoPagoHoras_R = lista_NoPagoHoras_R & "," & rsconsult!confval
                        End If
                    Case Else:
                        Flog.writeline Espacios(Tabulador * 1) & "***************************************************************"
                        Flog.writeline Espacios(Tabulador * 1) & "Error. Tipo TH. El nro columna '" & rsconsult!confnrocol & "' no es valido."
                        Flog.writeline Espacios(Tabulador * 1) & "***************************************************************"
                        GoTo Fin_error
                End Select
            Case Else:
                Flog.writeline Espacios(Tabulador * 1) & "***************************************************************"
                Flog.writeline Espacios(Tabulador * 1) & "Error. Tipo '" & rsconsult!conftipo & "' no reconocido en la configuración."
                Flog.writeline Espacios(Tabulador * 1) & "***************************************************************"
                GoTo Fin_error
        End Select
        rsconsult.MoveNext
    Loop
    rsconsult.Close
    
    
    '-------------------------------------------------------------------------
    'Comiensa a procesar
    '-------------------------------------------------------------------------
    Progreso = 0
    cantidadProcesada = DateDiff("d", fecdesde, fechasta) + 1
    If cantidadProcesada = 0 Then
        cantidadProcesada = 1
    End If
    IncPorc = (99 / cantidadProcesada)
    
    fecCalculo = fecdesde
    Do While fecCalculo <= fechasta
        
        Flog.writeline "** Día --> " & fecCalculo
        
        ' Busco los empleados que respetan el filtro inicial y posean datos en gti_acumdiario a la fecCalculo
        ' ordenados por los niveles organizacionales
        StrSql = "SELECT DISTINCT empleado.ternro "
        If tenro1 <> 0 Then
            StrSql = StrSql & ",his1.estrnro estrnro1,est1.estrdabr estrdabr1"
        End If
        If tenro2 <> 0 Then
            StrSql = StrSql & ",his2.estrnro estrnro2,est2.estrdabr estrdabr2"
        End If
        If tenro3 <> 0 Then
            StrSql = StrSql & ",his3.estrnro estrnro3,est3.estrdabr estrdabr3"
        End If
        StrSql = StrSql & " FROM empleado "
        StrSql = StrSql & " INNER JOIN his_estructura emp ON empleado.ternro = emp.ternro AND emp.tenro = 10"
        StrSql = StrSql & " AND (emp.htetdesde<=" & ConvFecha(fecCalculo) & " AND (emp.htethasta IS NULL OR emp.htethasta>=" & ConvFecha(fecCalculo) & "))"
        StrSql = StrSql & " AND emp.estrnro = " & empresa
        StrSql = StrSql & " INNER JOIN gti_acumdiario ON empleado.ternro = gti_acumdiario.ternro AND gti_acumdiario.adfecha = " & ConvFecha(fecCalculo)
        If tenro1 <> 0 Then
            StrSql = StrSql & " INNER JOIN his_estructura his1 ON empleado.ternro = his1.ternro AND his1.tenro = " & tenro1
            StrSql = StrSql & " AND (his1.htetdesde<=" & ConvFecha(fecCalculo) & " AND (his1.htethasta IS NULL OR his1.htethasta>=" & ConvFecha(fecCalculo) & "))"
            If estrnro1 <> -1 Then
                StrSql = StrSql & " AND his1.estrnro = " & estrnro1
            End If
            StrSql = StrSql & " INNER JOIN estructura est1 ON his1.estrnro = est1.estrnro "
        End If
        If tenro2 <> 0 Then
            StrSql = StrSql & " INNER JOIN his_estructura his2 ON empleado.ternro = his2.ternro AND his2.tenro = " & tenro2
            StrSql = StrSql & " AND (his2.htetdesde<=" & ConvFecha(fecCalculo) & " AND (his2.htethasta IS NULL OR his2.htethasta>=" & ConvFecha(fecCalculo) & "))"
            If estrnro2 <> -1 Then
                StrSql = StrSql & " AND his2.estrnro = " & estrnro2
            End If
            StrSql = StrSql & " INNER JOIN estructura est2 ON his2.estrnro = est2.estrnro "
        End If
        If tenro3 <> 0 Then
            StrSql = StrSql & " INNER JOIN his_estructura his3 ON empleado.ternro = his3.ternro AND his3.tenro = " & tenro3
            StrSql = StrSql & " AND (his3.htetdesde<=" & ConvFecha(fecCalculo) & " AND (his3.htethasta IS NULL OR his3.htethasta>=" & ConvFecha(fecCalculo) & "))"
            If estrnro3 <> -1 Then
                StrSql = StrSql & " AND his3.estrnro = " & estrnro3
            End If
            StrSql = StrSql & " INNER JOIN estructura est3 ON his3.estrnro = est3.estrnro "
        End If
        StrSql = StrSql & " WHERE empleg >= " & desde & " AND empleg <= " & hasta
        If estado <> 1 Then
            StrSql = StrSql & " AND empest = " & estado
        End If
        StrSql = StrSql & " ORDER BY "
        If tenro1 <> 0 Then
            StrSql = StrSql & "estrnro1, estrdabr1,"
        End If
        If tenro2 <> 0 Then
            StrSql = StrSql & "estrnro2, estrdabr2,"
        End If
        If tenro3 <> 0 Then
            StrSql = StrSql & "estrnro3, estrdabr3,"
        End If
        StrSql = StrSql & "empleado.ternro"
        
        OpenRecordset StrSql, rsconsult
        
        dias_Hab = 0
        hsNormal = 0
        resvpuesto = 0
        hspreact = 0
        pgohs = 0
        nopgohs = 0
        tothsaus = 0
        pgohsporc = 0
        nopgohsporc = 0
        hspres = 0
        hspresporc = 0
        hsextras = 0
        hsextrasporc = 0
        toths = 0
        tothsporc = 0
        cantemp = 0
        
        If rsconsult.EOF Then
        
            Flog.writeline "     No se encontraron Acumulados Diarios."
            Flog.writeline "      SQL --> " & StrSql
        
        Else
            
            contarDiasHab = True
            
            Do Until rsconsult.EOF
            
                'Determino si es un dia habil
                esFeriado = False
                If (Weekday(fecCalculo) = 7 Or Weekday(fecCalculo) = 1 Or objFeriado.Feriado(fecCalculo, rsconsult!Ternro, False)) Then
                    esFeriado = True
                Else
                
                    cantemp = cantemp + 1
                
                    If contarDiasHab Then
                        dias_Hab = dias_Hab + 1
                        contarDiasHab = False
                    End If
                End If
                
                'Horas Extras
                StrSql = "SELECT adcanthoras FROM gti_acumdiario "
                StrSql = StrSql & " WHERE thnro IN (" & lista_HsExtras & ") AND adfecha = " & ConvFecha(fecCalculo)
                StrSql = StrSql & " AND ternro = " & rsconsult!Ternro
                OpenRecordset StrSql, rs
                Do Until rs.EOF
                    hsextras = hsextras + rs!adcanthoras
                    rs.MoveNext
                Loop
                rs.Close
                
                If Not esFeriado Then
                    'Reserva Puesto
                    StrSql = "SELECT adcanthoras FROM gti_acumdiario "
                    StrSql = StrSql & " WHERE thnro IN (" & lista_ResPuesto & ") AND adfecha = " & ConvFecha(fecCalculo)
                    StrSql = StrSql & " AND ternro = " & rsconsult!Ternro
                    OpenRecordset StrSql, rs
                    Do Until rs.EOF
                        resvpuesto = resvpuesto + (rs!adcanthoras - CoefResta)
                        rs.MoveNext
                    Loop
                    rs.Close
                
                    'Pago de Horas
                    StrSql = "SELECT adcanthoras FROM gti_acumdiario "
                    StrSql = StrSql & " WHERE thnro IN (" & lista_PagoHoras & ") AND adfecha = " & ConvFecha(fecCalculo)
                    StrSql = StrSql & " AND ternro = " & rsconsult!Ternro
                    OpenRecordset StrSql, rs
                    Do Until rs.EOF
                        pgohs = pgohs + (rs!adcanthoras - CoefResta)
                        rs.MoveNext
                    Loop
                    rs.Close
                
                    'No Pago de Horas - Sumar
                    StrSql = "SELECT adcanthoras FROM gti_acumdiario "
                    StrSql = StrSql & " WHERE thnro IN (" & lista_NoPagoHoras_S & ") AND adfecha = " & ConvFecha(fecCalculo)
                    StrSql = StrSql & " AND ternro = " & rsconsult!Ternro
                    OpenRecordset StrSql, rs
                    Do Until rs.EOF
                        nopgohs = nopgohs + rs!adcanthoras
                        rs.MoveNext
                    Loop
                    rs.Close
                
                    'No Pago de Horas - Resta
                    StrSql = "SELECT adcanthoras FROM gti_acumdiario "
                    StrSql = StrSql & " WHERE thnro IN (" & lista_NoPagoHoras_R & ") AND adfecha = " & ConvFecha(fecCalculo)
                    StrSql = StrSql & " AND ternro = " & rsconsult!Ternro
                    OpenRecordset StrSql, rs
                    Do Until rs.EOF
                        nopgohs = nopgohs + Abs(rs!adcanthoras - CoefResta)
                        rs.MoveNext
                    Loop
                    rs.Close
                End If
                
                If tenro1 <> 0 Then
                    estrnro1_ant = rsconsult!estrnro1
                    estrdabr1_ant = rsconsult!estrdabr1
                End If
                If tenro2 <> 0 Then
                    estrnro2_ant = rsconsult!estrnro2
                    estrdabr2_ant = rsconsult!estrdabr2
                End If
                If tenro3 <> 0 Then
                    estrnro3_ant = rsconsult!estrnro3
                    estrdabr3_ant = rsconsult!estrdabr3
                End If
                
                rsconsult.MoveNext
                
                guardar = False
                If rsconsult.EOF Then
                    guardar = True
                Else
                    If tenro1 <> 0 Then
                        If estrdabr1_ant <> rsconsult!estrdabr1 Then
                            guardar = True
                        End If
                    End If
                    
                    If tenro2 <> 0 Then
                        If estrdabr2_ant <> rsconsult!estrdabr2 Then
                            guardar = True
                        End If
                    End If
                
                    If tenro3 <> 0 Then
                        If estrdabr3_ant <> rsconsult!estrdabr3 Then
                            guardar = True
                        End If
                    End If
                End If
                
                If guardar Then
                    ' Busco si se encuentran valores para los niveles organizacionales
                    ' Si no encuentro --> inserto
                    ' Si encuentro valores
                    '       Verifico si el ultimo dia (fhasta) es imnediatamente anterior al fecCalculo y la dotacion de personal coincide
                    '           Recalculo los valores
                    '           fhasta = fecCalculo
                    '       Sino
                    '           inserto un nuevo valor con nuevo intervalo de fecha y dotacion de personal
                    
                    StrSql = "SELECT * FROM rep_mano_obra_det WHERE bpronro = " & NroProceso
                    If tenro1 <> 0 Then
                        StrSql = StrSql & " AND estrnro1 = " & estrnro1_ant
                    End If
                    If tenro2 <> 0 Then
                        StrSql = StrSql & " AND estrnro2 = " & estrnro2_ant
                    End If
                    If tenro3 <> 0 Then
                        StrSql = StrSql & " AND estrnro3 = " & estrnro3_ant
                    End If
                    OpenRecordset StrSql, rs
                    
                    If rs.EOF Then
                        'Insertar
                        insertarRegistro = True
                    Else
                        encontre_rango = True
                        Do While (encontre_rango)
                            If rs.EOF Then
                                encontre_rango = False
                            Else
                                If (rs!fhasta = DateAdd("d", -1, fecCalculo) And (rs!dotacion = cantemp Or esFeriado)) Then
                                    encontre_rango = False
                                Else
                                    rs.MoveNext
                                End If
                            End If
                        Loop
                        
                        If rs.EOF Then
                            ' Insertar
                            insertarRegistro = True
                        Else
                            ' Recalcular
                            insertarRegistro = False
                        End If
                    End If
                    
                    
                    If insertarRegistro Then
                        'Insertar registro
                        hsNormal = (cantemp * CoefHsNormal * dias_Hab)
                        hspreact = ((cantemp * CoefHsPeAct * dias_Hab) - resvpuesto)
                        tothsaus = pgohs + nopgohs
                        If tothsaus <> 0 Then
                            pgohsporc = (pgohs * 100) / tothsaus
                            nopgohsporc = (nopgohs * 100) / tothsaus
                        End If
                        hspres = hspreact - tothsaus
                        If hsNormal <> 0 Then
                            hspresporc = (hspres * 100) / hsNormal
                        End If
                        If hspres <> 0 Then
                            hsextrasporc = (hsextras * 100) / hspres
                        End If
                        toths = hspres + hsextras
                        If hspres <> 0 Then
                            tothsporc = (toths * 100) / hspres
                        End If
                        'Inserto
                        StrSql = "INSERT INTO rep_mano_obra_det (bpronro,estrnro1,estrdabr1,estrnro2,estrdabr2,estrnro3," & _
                            "estrdabr3,fdesde,fhasta,diasHab,dotacion,hsnormal,resvpuesto,hspreact,pgohs,pgohsporc,nopgohs," & _
                            "nopgohsporc,tothsaus,hspres,hspresporc,hsextras,hsextrasporc,toths,tothsporc) VALUES (" & _
                            NroProceso & "," & estrnro1_ant & ",'" & estrdabr1_ant & "'," & estrnro2_ant & ",'" & _
                            estrdabr2_ant & "'," & estrnro3_ant & ",'" & estrdabr3_ant & "'," & _
                            ConvFecha(fecCalculo) & "," & ConvFecha(fecCalculo) & "," & dias_Hab & "," & cantemp & "," & _
                            hsNormal & "," & resvpuesto & "," & hspreact & "," & pgohs & "," & pgohsporc & "," & nopgohs & "," & _
                            nopgohsporc & "," & tothsaus & "," & hspres & "," & hspresporc & "," & _
                            hsextras & "," & hsextrasporc & "," & toths & "," & tothsporc & ")"
                        objConn.Execute StrSql, , adExecuteNoRecords
                        
                        Flog.writeline "     Inserto Registro."
                        Flog.writeline "      SQL --> " & StrSql
                        
                    Else
                        'Update registro
                        dias_Hab = dias_Hab + rs!diasHab
                        hsNormal = (rs!dotacion * CoefHsNormal * dias_Hab)
                        resvpuesto = resvpuesto + rs!resvpuesto
                        hspreact = ((rs!dotacion * CoefHsPeAct * dias_Hab) - resvpuesto)
                        pgohs = pgohs + rs!pgohs
                        nopgohs = nopgohs + rs!nopgohs
                        tothsaus = pgohs + nopgohs
                        If tothsaus <> 0 Then
                            pgohsporc = (pgohs * 100) / tothsaus
                            nopgohsporc = (nopgohs * 100) / tothsaus
                        End If
                        hspres = hspreact - tothsaus
                        If hsNormal <> 0 Then
                            hspresporc = (hspres * 100) / hsNormal
                        End If
                        hsextras = hsextras + rs!hsextras
                        If hspres <> 0 Then
                            hsextrasporc = (hsextras * 100) / hspres
                        End If
                        toths = hspres + hsextras
                        If hspres <> 0 Then
                            tothsporc = (toths * 100) / hspres
                        End If
                        'Update
                        StrSql = "UPDATE rep_mano_obra_det SET " & _
                            "hsnormal = " & hsNormal & ",fhasta = " & ConvFecha(fecCalculo) & ",diasHab=" & dias_Hab & _
                            ",resvpuesto = " & resvpuesto & ",hspreact = " & hspreact & ",pgohs = " & pgohs & _
                            ",pgohsporc = " & pgohsporc & ",nopgohs = " & nopgohs & _
                            ",nopgohsporc = " & nopgohsporc & ",tothsaus = " & tothsaus & ",hspres = " & hspres & _
                            ",hspresporc = " & hspresporc & ",hsextras = " & hsextras & _
                            ",hsextrasporc = " & hsextrasporc & ",toths = " & toths & _
                            ",tothsporc = " & tothsporc & _
                            " WHERE bpronro = " & NroProceso & " AND estrnro1 = " & rs!estrnro1 & _
                            " AND estrnro2 = " & rs!estrnro2 & " AND estrnro3 = " & rs!estrnro3 & _
                            " AND fdesde = " & ConvFecha(rs!fdesde) & " AND fhasta = " & ConvFecha(rs!fhasta)
                        objConn.Execute StrSql, , adExecuteNoRecords
                    
                        Flog.writeline "     Update Registro."
                        Flog.writeline "      SQL --> " & StrSql
                        
                    End If
                    rs.Close
                                            
                    'Inicialiso los valores
                    'dias_Hab = 0
                    hsNormal = 0
                    resvpuesto = 0
                    hspreact = 0
                    pgohs = 0
                    nopgohs = 0
                    tothsaus = 0
                    pgohsporc = 0
                    nopgohsporc = 0
                    hspres = 0
                    hspresporc = 0
                    hsextras = 0
                    hsextrasporc = 0
                    toths = 0
                    tothsporc = 0
                    cantemp = 0
                End If
            Loop
                
                
        End If
        
        rsconsult.Close
        
        'Actualizo el progreso
        Progreso = Progreso + IncPorc
        cantidadProcesada = cantidadProcesada - 1
        
        TiempoAcumulado = GetTickCount
        StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso
        StrSql = StrSql & ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'"
        StrSql = StrSql & ", bprcempleados ='" & CStr(cantidadProcesada) & "' WHERE bpronro = " & NroProceso
        objConnProgreso.Execute StrSql, , adExecuteNoRecords
        
        fecCalculo = DateAdd("d", 1, fecCalculo)
        
    Loop
    
    Flog.writeline " "
    
    MyCommitTrans
    
Fin:
Exit Sub
            
Fin_error:
    MyRollbackTrans
    Exit Sub

MError:
    MyRollbackTrans
'    Resume Next
    Flog.writeline
    Flog.writeline "***************************************************************"
    Flog.writeline " Error: " & Err.Description
    Flog.writeline " Última Sql ejecutada: " & StrSql
    Flog.writeline "***************************************************************"
    Flog.writeline
    HuboErrores = True
    Exit Sub
End Sub
            
