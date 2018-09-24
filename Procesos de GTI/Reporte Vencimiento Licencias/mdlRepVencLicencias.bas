Attribute VB_Name = "mdlRepVencLicencias"
Option Explicit

Const Version = "1.00"
Const FechaVersion = "07/08/2015"
'Modificaciones: Sebastian Stremel - CAS-30798 - ACARA - SO - Nuevo reporte Historia clínica
'

Global HuboErrores As Boolean
Global Usuario As String
Global cadena As String
Dim dirsalidas As String
Dim diasAntes As Integer
Dim cantAnios
Dim fecha As String




Sub Main()
Dim Archivo As String
Dim pos As Integer
Dim strcmdLine As String

'Dim objconnMain As New ADODB.Connection
Dim rs As New ADODB.Recordset

Dim fecha As Date
Dim Hora As String
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
    Archivo = PathFLog & "RepVencLicencias-" & CStr(NroProceso) & Format(Now, "DD-MM-YYYY") & ".log"
    
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
    
    'Cambio el estado del proceso a Procesando
    StrSql = "UPDATE batch_proceso SET bprcpid = " & PID & ", bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcestado = 'Procesando', bprcprogreso = 0 WHERE bpronro = " & NroProceso
    objConn.Execute StrSql, , adExecuteNoRecords

    Flog.writeline Espacios(Tabulador * 0) & "Levanta Proceso y Setea Parámetros:  " & " " & Now
    
    'levanto los parametros del proceso
    StrParametros = ""
    StrSql = "SELECT bprcfecdesde,bprcfechasta,bprcparam,bprcfecha,bprchora,iduser  "
    StrSql = StrSql & " FROM batch_proceso "
    StrSql = StrSql & " WHERE bpronro = " & NroProceso
    OpenRecordset StrSql, rs
    If Not rs.EOF Then
        fecha = rs!bprcfecha
        Hora = rs!bprchora
        Usuario = rs!iduser
        If Not IsNull(rs!bprcparam) Then
            diasAntes = rs!bprcparam
            Flog.writeline "El proceso avisara " & diasAntes & " del vencimiento de las licencias"
        End If
    Else
        Exit Sub
    End If
    

    Call repVencimiento(NroReporte, NroProceso, diasAntes, fecha, Hora)
    
    
    'Actualizo el Btach_Proceso
    If Not HuboErrores Then
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado' WHERE bpronro = " & NroProceso
    Else
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Incompleto' WHERE bpronro = " & NroProceso
    End If
    objConn.Execute StrSql, , adExecuteNoRecords
    
    'Cierro y libero todo
    If TransactionRunning Then MyRollbackTrans
    
    If objConn.State = adStateOpen Then objConn.Close
    If objConnProgreso.State = adStateOpen Then objConnProgreso.Close
    If CnTraza.State = adStateOpen Then CnTraza.Close
Exit Sub

ce:
    Flog.writeline "Reporte abortado por Error:" & " " & Now
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 0) & "**********************************************************"
    Flog.writeline Espacios(Tabulador * 0) & " Error General " & Err.Description
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 0) & "SQL Ejecutado: " & StrSql
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 0) & "**********************************************************"
    Flog.writeline
        
    'Actualizo el progreso
    MyBeginTrans
        StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcprogreso = 0, bprcestado = 'Error General', bprcpid = " & PID & " WHERE bpronro = " & NroProceso
        objConnProgreso.Execute StrSql, , adExecuteNoRecords
    MyCommitTrans
    
    
End Sub

Public Sub repVencimiento(ByVal NroReporte As Long, ByVal NroProceso As Long, ByVal diasAntes As Integer, ByVal fecha As Date, ByVal Hora As String)

Dim rs As New ADODB.Recordset
Dim listaEmpleados As String
Dim empleados
Dim empleado As Integer
Dim j As Integer
Dim str_licencias As String

j = 0
listaEmpleados = 0

'________________________________________________
'LEVANTO LOS DATOS DEL CONFREP
StrSql = " SELECT confnrocol, confval, confval2 FROM confrep "
StrSql = StrSql & " WHERE repnro=492"
OpenRecordset StrSql, rs
If Not rs.EOF Then
    Do While Not rs.EOF
        Select Case rs("confnrocol")
            Case 2:
                If Not IsNull(rs("confval")) Then
                    cantAnios = rs("confval")
                Else
                    cantAnios = 2
                End If
                
        End Select
    rs.MoveNext
    Loop
End If
rs.Close
cantAnios = cantAnios * -1
'________________________________________________

'________________________________________________
'BUSCO LOS EMPLEADOS QUE TIENE CARGADAS LICENCIAS
StrSql = " SELECT distinct(empleado.empleg) legajo, ternro "
StrSql = StrSql & " From emp_lic "
StrSql = StrSql & " INNER JOIN empleado ON empleado.ternro = emp_lic.empleado "
StrSql = StrSql & " ORDER BY empleado.empleg ASC "
OpenRecordset StrSql, rs
If Not rs.EOF Then
    Do While Not rs.EOF
        listaEmpleados = listaEmpleados & "," & rs!ternro
    rs.MoveNext
    Loop
Else
    Flog.writeline "No hay empleados para analizar."
    Exit Sub
End If
rs.Close
'________________________________________________


empleados = Split(listaEmpleados, ",")
cadena = ""
For j = 1 To UBound(empleados)
    'If empleados(j) = 2049 Or empleados(j) = 2050 Then
        empleado = empleados(j)
        '____________________________
        StrSql = "SELECT sovisitamedica.vismednro,empleado.empleg, empleado.terape, empleado.ternom, emp_lic.elfechadesde, emp_lic.elfechahasta, emp_lic.elcantdias,  sovisitamedica.vismeddiag "
        StrSql = StrSql & ",servmedico, emp_lic.tdnro, emp_lic.empleado, emp_lic.emp_licnro,sovisitamedica.vismeddesc, sopatol_visitas.patologianro"
        StrSql = StrSql & " FROM emp_lic "
        StrSql = StrSql & " INNER JOIN empleado ON empleado.ternro = emp_lic.empleado "
        StrSql = StrSql & " INNER JOIN licencia_visita ON licencia_visita.emp_licnro = emp_lic.emp_licnro "
        StrSql = StrSql & " INNER JOIN sovisitamedica ON sovisitamedica.vismednro = licencia_visita.visitamed "
        StrSql = StrSql & " INNER JOIN sopatol_visitas ON sopatol_visitas.visitamed = sovisitamedica.vismednro "
        StrSql = StrSql & " WHERE (emp_lic.empleado = " & empleado & ") "
        StrSql = StrSql & " ORDER BY elfechadesde ASC"
        OpenRecordset StrSql, rs
        If Not rs.EOF Then
            Flog.writeline "Empieza a calcular el Vto del empleado: " & rs("empleado")
            Do While Not rs.EOF
                'para cada licencia calculo el vencimiento
                
                calcularVto rs("tdnro"), rs("empleado"), rs("elfechadesde"), rs("elfechahasta"), rs("patologianro")
            rs.MoveNext
            Loop
        End If
    'End If
    '____________________________
Next j

If cadena <> "" Then
    str_licencias = "<table>"
    str_licencias = str_licencias & "<tr>"
    str_licencias = str_licencias & "<th colspan=""6"">Vencimiento de licencias</th>"
    str_licencias = str_licencias & "</tr>"
    
    str_licencias = str_licencias & "<tr>"
    str_licencias = str_licencias & "<th>Legajo</th>"
    str_licencias = str_licencias & "<th>nombre y apellido</th>"
    str_licencias = str_licencias & "<th>licencia</th>"
    str_licencias = str_licencias & "<th>patología</th>"
    str_licencias = str_licencias & "<th>fecha de inicio</th>"
    str_licencias = str_licencias & "<th>fecha de vencimiento</th>"
    str_licencias = str_licencias & "</tr>"
    
    str_licencias = str_licencias & cadena
    str_licencias = str_licencias & "</table>"
    Flog.writeline "Cadena: " & cadena
    

    
    'disparar procesos de mensajeria
    Call crearProcesosMensajeria(str_licencias, "")
End If


End Sub
Sub crearProcesosMensajeria(ByRef str_error As String, ByVal empresa As String)

Dim objRs As New ADODB.Recordset
Dim fs2, MsgFile
Dim titulo As String
Dim bpronroMail As Long
Dim mails As String
Dim notiNro As Long
Dim mailFileName As String
Dim mailFile


    ' Directorio Salidas
    StrSql = "SELECT sis_dirsalidas FROM sistema"
    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then
        dirsalidas = objRs!sis_dirsalidas & "\attach"
        Flog.writeline "Directorio de Salidas: " & dirsalidas
    Else
        Flog.writeline "No se encuentra configurado sis_dirsalidas"
        Exit Sub
    End If
    If objRs.State = adStateOpen Then objRs.Close

    'Busco el codigo de la notificacion
    StrSql = "SELECT confval FROM confrep "
    StrSql = StrSql & " WHERE conftipo = 'TN' AND repnro = 492"
    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then
        notiNro = objRs!confval
    Else
        Flog.writeline Espacios(Tabulador * 0) & "No esta configurado el tipo de alerta, para el envio de mail."
        notiNro = 0
    End If
    
    If objRs.State = adStateOpen Then objRs.Close
    'FGZ - 04/09/2006 - Saco esto y lo pongo afuera
    StrSql = "insert into batch_proceso "
    StrSql = StrSql & "(btprcnro, bprcfecha, iduser, bprchora, bprcfecdesde, bprcfechasta, bprcparam, "
    StrSql = StrSql & "bprcestado, bprcprogreso, bprcfecfin, bprchorafin, bprctiempo, empnro, bprcempleados) "
    StrSql = StrSql & "values (25," & ConvFecha(Date) & ",'" & Usuario & "','" & FormatDateTime(Time, 4) & ":00'"
    StrSql = StrSql & ",null,null,'1','Pendiente',null,null,null,null,0,null)"
    
    objConn.Execute StrSql, , adExecuteNoRecords
    
    bpronroMail = getLastIdentity(objConn, "batch_proceso")


    '--------------------------------------------------
    'Busco todos los usuarios a los cuales les tengo que enviar los mails
    StrSql = "SELECT usremail FROM user_per "
    StrSql = StrSql & "inner join noti_usuario on user_per.iduser = noti_usuario.iduser "
    StrSql = StrSql & "where notinro = " & notiNro
    OpenRecordset StrSql, objRs
    mails = ""
    Do Until objRs.EOF
        If Not IsNull(objRs!usremail) Then
            If Len(objRs!usremail) > 0 Then
                mails = mails & objRs!usremail & ";"
            End If
        End If
        objRs.MoveNext
    Loop
    
    mailFileName = dirsalidas & "\msg_" & bpronroMail & "_Vencimiento de Licencias_" & Replace(FormatDateTime(Date, 2), "/", "-") & "_" & Replace(FormatDateTime(Time, 4), ":", "-") & "-" & String(2 - Len(Second(Now)), "0") & Second(Now)
    Set fs2 = CreateObject("Scripting.FileSystemObject")
    Set mailFile = fs2.CreateTextFile(mailFileName & ".html", True)
    
    mailFile.writeline "<html><head>"
    mailFile.writeline "<title> Licencias con Vencimientos - RHPro &reg; </title></head><body>"
    'mailFile.writeline "<h4>Errores Detectados</h4>"
    mailFile.writeline "<table>" & str_error & "</table>"
    mailFile.writeline "</body></html>"
    mailFile.Close
    '--------------------------------------------------


    Set fs2 = CreateObject("Scripting.FileSystemObject")
    
    'FGZ - 05/09/2006
    'Los nombres de los archivos para los mails de esta alerta empiezan con el bpronro de este proceso
    Set MsgFile = fs2.CreateTextFile(mailFileName & ".msg", True)
    
    MsgFile.writeline "[MailMessage]"
    MsgFile.writeline "FromName=RHPro - Vencimiento de Licencias"
    MsgFile.writeline "Subject=Vencimientos de licencias "
    MsgFile.writeline "Body1=Vencimiento de licencias"
    If Len(mailFileName) > 0 Then
       MsgFile.writeline "Attachment=" & mailFileName & ".html"
    Else
       MsgFile.writeline "Attachment="
    End If
    
    MsgFile.writeline "Recipients=" & mails
    
    
    If objRs.State = adStateOpen Then objRs.Close
    
    StrSql = "select cfgemailfrom,cfgemailhost,cfgemailport,cfgemailuser,cfgemailpassword,cfgssl from conf_email where cfgemailest = -1"
    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then
        MsgFile.writeline "FromAddress=" & objRs!cfgemailfrom
        MsgFile.writeline "Host=" & objRs!cfgemailhost
        MsgFile.writeline "Port=" & objRs!cfgemailport
        MsgFile.writeline "User=" & objRs!cfgemailuser
        MsgFile.writeline "Password=" & objRs!cfgemailpassword
    Else
        Flog.writeline "No existen datos configurados para el envio de emails, o no existe configuracion activa"
        Exit Sub
    End If
    MsgFile.writeline "CCO="
    MsgFile.writeline "CC="
    MsgFile.writeline "HTMLBody="
    MsgFile.writeline "HTMLMailHeader="
    MsgFile.writeline "SSL=" & objRs!cfgssl
    
    If objRs.State = adStateOpen Then objRs.Close

End Sub
Function tieneFamiliares(ByVal ternro As Long) As Boolean
Dim StrSql As String
Dim rs As New ADODB.Recordset

    StrSql = "SELECT tercero.ternro "
    StrSql = StrSql & " FROM  tercero "
    StrSql = StrSql & " INNER JOIN familiar ON tercero.ternro=familiar.ternro "
    StrSql = StrSql & " WHERE familiar.empleado = " & ternro
    OpenRecordset StrSql, rs
    If Not rs.EOF Then
        tieneFamiliares = True
    Else
        tieneFamiliares = False
    End If
    rs.Close
    
End Function

Function buscarFalta(ByVal ternro As Long)
Dim StrSql As String
Dim rs_aux As New ADODB.Recordset
    StrSql = " select MIN(altfec) falta from fases where empleado =" & ternro
    OpenRecordset StrSql, rs_aux
    If Not rs_aux.EOF Then
        buscarFalta = rs_aux("falta")
    End If
    rs_aux.Close
End Function

Public Sub calcularVto(ByVal tipdia As Integer, ByVal empleado As Long, ByVal fechadesde As Date, ByVal fechahasta As Date, ByVal patologia As Integer)

Dim cant As Integer

Dim cantAux As Integer
Dim total As Integer
Dim falta As String
Dim hayFamiliares As Boolean

Dim rs As New ADODB.Recordset
Dim empleg As Long
Dim empApe As String
Dim empNom As String
Dim patDesc As String
Dim licDesc As String

hayFamiliares = tieneFamiliares(empleado)

cant = buscarDiasTomados(tipdia, empleado, fechadesde, fechahasta, patologia) 'dias tomados hasta el momento
falta = buscarFalta(empleado)
    
'busco el legajo del empleado
StrSql = " SELECT empleg, terape, ternom, terape2, ternom2 "
StrSql = StrSql & " FROM empleado "
StrSql = StrSql & " WHERE ternro=" & empleado
OpenRecordset StrSql, rs
If Not rs.EOF Then
    empleg = rs!empleg
    empApe = IIf(EsNulo(rs!terape2), rs!terape, rs!terape & " " & rs!terape2)
    empNom = IIf(EsNulo(rs!ternom2), rs!ternom, rs!ternom & " " & rs!ternom2 & ", ")
End If
rs.Close
'hasta aca
    
'Busco la descripcion de la patologia
StrSql = " SELECT patologiadesabr FROM sopatologias "
StrSql = StrSql & " WHERE patologianro=" & patologia
OpenRecordset StrSql, rs
If Not rs.EOF Then
    patDesc = IIf(EsNulo(rs!patologiadesabr), "", rs!patologiadesabr)
End If
rs.Close
'hasta aca

'Busco la descripcion del tipo de lic
StrSql = " SELECT tddesc FROM tipdia "
StrSql = StrSql & " WHERE tdnro=" & tipdia
OpenRecordset StrSql, rs
If Not rs.EOF Then
    licDesc = IIf(EsNulo(rs!tddesc), "", rs!tddesc)
End If
rs.Close
'hasta aca
    
    Select Case tipdia
        Case 8
            cantAux = 90 - cant
            total = cant + DateDiff("d", fechadesde, fechahasta) + 1
            If total >= 90 Then 'llego a los 90 dias
                If cantAux < 0 Then
                    fecha = fecha
                Else
                    fecha = DateAdd("d", (90 - total), fechahasta)
                End If
                If DateDiff("YYYY", fecha, falta) <= 5 And Not hayFamiliares Then
                    'VERIFICO SI FALTAN X DIAS PARA EL VENC
                    If (CDate(DateAdd("d", diasAntes, Date)) = CDate(fecha)) Or ((DateDiff("d", fecha, Date) <= diasAntes And (DateDiff("d", fecha, Date) >= 0)) Or (DateDiff("d", fecha, Date) >= (diasAntes * -1)) And (DateDiff("d", fecha, Date) <= 0)) Then
                        cadena = cadena & "<tr>"
                        cadena = cadena & "<td>" & empleg & "</td>"
                        cadena = cadena & "<td>" & empNom & empApe & "</td>"
                        cadena = cadena & "<td>" & licDesc & "</td>"
                        cadena = cadena & "<td>" & patDesc & "</td>"
                        cadena = cadena & "<td>" & fechadesde & "</td>"
                        cadena = cadena & "<td>" & fecha & "</td>"
                        cadena = cadena & "</tr>"
                        'response.Write fecha
                    End If
                    'HASTA ACA
                End If
            Else
                cantAux = 180 - cant
                total = cant + DateDiff("d", fechadesde, fechahasta) + 1
                If total >= 180 Then 'llego a los 180 dias (6 meses)
                    If cantAux < 0 Then
                        fecha = fecha
                    Else
                        fecha = DateAdd("d", (180 - total), fechahasta)
                    End If
                    
                    If (DateDiff("YYYY", fecha, falta) <= 5 And hayFamiliares) Or (DateDiff("YYYY", fecha, falta) > 5 And Not hayFamiliares) Then
                        'VERIFICO SI FALTAN X DIAS PARA EL VENC
                        If (CDate(DateAdd("d", diasAntes, Date)) = CDate(fecha)) Or ((DateDiff("d", fecha, Date) <= 10 And (DateDiff("d", fecha, Date) >= 0)) Or (DateDiff("d", fecha, Date) >= -10) And (DateDiff("d", fecha, Date) <= 0)) Then
                            cadena = cadena & "<tr>"
                            cadena = cadena & "<td>" & empleg & "</td>"
                            cadena = cadena & "<td>" & empNom & empApe & "</td>"
                            cadena = cadena & "<td>" & licDesc & "</td>"
                            cadena = cadena & "<td>" & patDesc & "</td>"
                            cadena = cadena & "<td>" & fechadesde & "</td>"
                            cadena = cadena & "<td>" & fecha & "</td>"
                            cadena = cadena & "</tr>"
                        End If
                        'HASTA ACA
                    End If
                Else
                    cantAux = 365 - cant
                    total = cant + DateDiff("d", fechadesde, fechahasta) + 1
                    If total >= 365 Then
                        If cantAux < 0 Then
                            fecha = fecha
                        Else
                            fecha = DateAdd("d", (365 - total), fechahasta)
                        End If
                        If (DateDiff("YYYY", fecha, falta) > 5 And hayFamiliares) Then
                            'VERIFICO SI FALTAN X DIAS PARA EL VENC
                            If (CDate(DateAdd("d", diasAntes, Date)) = CDate(fecha)) Or ((DateDiff("d", fecha, Date) <= 10 And (DateDiff("d", fecha, Date) >= 0)) Or (DateDiff("d", fecha, Date) >= -10) And (DateDiff("d", fecha, Date) <= 0)) Then
                                cadena = cadena & "<tr>"
                                cadena = cadena & "<td>" & empleg & "</td>"
                                cadena = cadena & "<td>" & empNom & empApe & "</td>"
                                cadena = cadena & "<td>" & licDesc & "</td>"
                                cadena = cadena & "<td>" & patDesc & "</td>"
                                cadena = cadena & "<td>" & fechadesde & "</td>"
                                cadena = cadena & "<td>" & fecha & "</td>"
                                cadena = cadena & "</tr>"
                            End If
                            'HASTA ACA
                        End If
                    End If
                End If
            End If
        Case 9, 13, 14
            cantAux = 365 - cant
            total = cant + DateDiff("d", fechadesde, fechahasta) + 1
            If total >= 365 Then
                If cantAux < 0 Then
                    fecha = fecha
                Else
                    fecha = DateAdd("d", (365 - total), fechahasta)
                End If
                'fecha = DateAdd("d", (365 - total), fechahasta)
                'VERIFICO SI FALTAN X DIAS PARA EL VENC
                If (CDate(DateAdd("d", diasAntes, Date)) = CDate(fecha)) Or ((DateDiff("d", fecha, Date) <= 10 And (DateDiff("d", fecha, Date) >= 0)) Or (DateDiff("d", fecha, Date) >= -10) And (DateDiff("d", fecha, Date) <= 0)) Then
                    cadena = cadena & "<tr>"
                    cadena = cadena & "<td>" & empleg & "</td>"
                    cadena = cadena & "<td>" & empNom & empApe & "</td>"
                    cadena = cadena & "<td>" & licDesc & "</td>"
                    cadena = cadena & "<td>" & patDesc & "</td>"
                    cadena = cadena & "<td>" & fechadesde & "</td>"
                    cadena = cadena & "<td>" & fecha & "</td>"
                    cadena = cadena & "</tr>"
                End If
                'HASTA ACA
                'response.Write fecha
            End If
    End Select
    
End Sub


Function buscarDiasTomados(tipdia, empleado, fechadesde, fechahasta, patologianro)

Dim diasA
Dim diasD
Dim cantidad
Dim l_fechadesdeAux
Dim l_fechadesde
Dim rs As New ADODB.Recordset
Dim StrSql As String



l_fechadesde = fechadesde
    If fechadesde <> "" Then
        l_fechadesdeAux = DateAdd("YYYY", cantAnios, fechadesde)
    End If
    StrSql = " SELECT  elfechadesde desde, elfechahasta hasta, elcantdias cant "
    StrSql = StrSql & " FROM emp_lic "
    StrSql = StrSql & " INNER JOIN v_empleado ON v_empleado.ternro = emp_lic.empleado "
    If tipdia <> 9 And tipdia <> 13 And tipdia <> 14 Then    'enfermedad
        StrSql = StrSql & " INNER JOIN licencia_visita ON licencia_visita.emp_licnro = emp_lic.emp_licnro "
        StrSql = StrSql & " INNER JOIN sopatol_visitas ON sopatol_visitas.visitamed = licencia_visita.visitamed "
        StrSql = StrSql & " WHERE emp_lic.empleado =" & empleado
        If patologianro <> "" Then
            StrSql = StrSql & " AND sopatol_visitas.patologianro = " & patologianro
        End If
    End If
    If tipdia = 9 Or tipdia = 13 Or tipdia = 14 Then 'accidente
        StrSql = StrSql & " INNER JOIN lic_accid lica ON lica.emp_licnro = emp_lic.emp_licnro "
        StrSql = StrSql & " WHERE emp_lic.empleado =" & empleado
        StrSql = StrSql & " AND emp_lic.tdnro IN (9, 13, 14)"
    Else
        StrSql = StrSql & " AND emp_lic.tdnro = " & tipdia
    End If

    If l_fechadesdeAux <> "" Then
        StrSql = StrSql & " AND ((elfechadesde <= " & ConvFecha(l_fechadesdeAux) & " AND (elfechahasta is null or elfechahasta >= " & ConvFecha(l_fechadesde)
        StrSql = StrSql & " or elfechahasta >= " & ConvFecha(l_fechadesdeAux) & ")) OR "
        StrSql = StrSql & " (elfechadesde >= " & ConvFecha(l_fechadesdeAux) & " AND (elfechadesde <= " & ConvFecha(l_fechadesde) & ")))"
    End If
    OpenRecordset StrSql, rs

    If Not rs.EOF Then
        Do While Not rs.EOF
            cantidad = 0
            If ((CDate(rs("desde")) < CDate(l_fechadesdeAux)) And (CDate(rs("hasta")) > CDate(l_fechadesde))) Then
                'response.write "A"
                diasA = DateDiff("d", rs("desde"), l_fechadesdeAux)
                diasA = diasA + 1
                diasD = DateDiff("d", l_fechadesde, rs("hasta"))
                diasD = diasD + 1
                cantidad = CDbl(rs("cant")) - (diasA + diasD)
                buscarDiasTomados = CDbl(buscarDiasTomados) + CDbl(cantidad)
            Else
                If rs("desde") < l_fechadesdeAux Then
                    'response.write "B"
                    diasA = DateDiff("d", rs("desde"), l_fechadesdeAux)
                    diasA = diasA + 1
                    cantidad = CDbl(rs("cant")) - (diasA)
                    buscarDiasTomados = CDbl(buscarDiasTomados) + CDbl(cantidad)
                Else
                    If (rs("hasta") > l_fechadesde) Then
                        'response.write "C"
                        diasD = DateDiff("d", l_fechadesde, rs("hasta"))
                        diasD = diasD + 1 ' le sumo uno porque datediff no tiene en cuenta los extremos
                        cantidad = CDbl(rs("cant")) - (diasD)
                        buscarDiasTomados = CDbl(buscarDiasTomados) + CDbl(cantidad)
                    Else
                        'response.write "D"
                        buscarDiasTomados = CDbl(buscarDiasTomados) + CDbl(rs("cant"))
                    End If
                End If
            End If
        rs.MoveNext
        Loop
    Else
        buscarDiasTomados = 0
    End If
    
End Function
