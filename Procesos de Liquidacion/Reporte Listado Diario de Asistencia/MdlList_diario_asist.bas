Attribute VB_Name = "MdlListado_diario_asist"
Option Explicit

'Const Version = 1
'Const FechaVersion = "23/05/2011"
'------------------------------------------------------------------------------------------
'Autor: Gonzalez Nicolás
'------------------------------------------------------------------------------------------
'Const Version = "1.01"
'Const FechaVersion = "04/07/2011"
'Modificado: Gonzalez Nicolás - Se cambío descripción de la licencia para las vacaciones
'                              - Se agregó función CHoras()
'Const Version = "1.02"
'Const FechaVersion = "07/07/2011"
'Modificado: Gonzalez Nicolás - Se cambío funcion tipodehora_acum()
'------------------------------------------------------------------------------------------
'Const Version = "1.03"
'Const FechaVersion = "08/09/2011"
'Modificado: Gonzalez Nicolás -
'                             - Se agregó X en lugar de cantidad de hs cuando no hay registraciones de un empleado
'------------------------------------------------------------------------------------------
'Const Version = "1.04"
'Const FechaVersion = "09/11/2011"
'Modificado: Manterola Maria Magdalena -
'                             - Se modificó la columna Falta, para que muestre una 'X' cuando no hay registraciones de un empleado, y que la muestre vacia en caso contrario.
'------------------------------------------------------------------------------------------
'Const Version = "1.05"
'Const FechaVersion = "26/12/2011"
'Modificado: Manterola Maria Magdalena -
'                             - Se modificó el reporte para que incluya solamente personal que tiene novedades, es decir no debería traer personas sin anormalidad para ese día.
'------------------------------------------------------------------------------------------
'Const Version = "1.06"
'Const FechaVersion = "29/12/2011"
'Modificado: Manterola Maria Magdalena -
'                             - Se agregó al buscar la novedad que busque aquellas que tengan jussigla = NOV
'------------------------------------------------------------------------------------------
'Const Version = "1.07"
'Const FechaVersion = "20/03/2012"
'Modificado: Carmen Quintero - (12834) - Se cambió v_empleado por empleado
'------------------------------------------------------------------------------------------

'Const Version = "1.08"
'Const FechaVersion = "22/03/2012"
'Modificado: Carmen Quintero - (12834) - Se modificó la funcion lic_anual para que considere el caso, cuando la suma de cantidad de dias venga vacia
'------------------------------------------------------------------------------------------

Const Version = "1.09"
Const FechaVersion = "08/05/2012"
'Modificado: Manterola Maria Magdalena - (15717) - Se modificó la funcion lic_anual para que considere el caso, cuando el limite por evento(tdinteger4) del tipo de licencia esta cargado en NULL o vacío. En dicho caso le asigna 0 a topeanual.

'------------------------------------------------------------------------------------------

Public Type TipoRestriccion
    Estrnro As Long
    Valor As Double
End Type

Global nListaProc As String
Global usuario As String








Public Sub Main()
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento inicial de Reporte Listado Diario de Asistencia.
' Autor      : Gonzalez Nicolás
' Fecha      : 23/05/2011
' Descripcion: Genera reporte para MUNDO MAIPU
' Ultima Mod.:
' ---------------------------------------------------------------------------------------------
Dim strCmdLine
Dim Nombre_Arch As String
Dim rs_batch_proceso As New ADODB.Recordset
Dim PID As String
Dim bprcparam As String
Dim ArrParametros

    strCmdLine = Command()
    ArrParametros = Split(strCmdLine, " ", -1)
    If UBound(ArrParametros) > 1 Then
        If IsNumeric(ArrParametros(0)) Then
            NroProcesoBatch = ArrParametros(0)
            Etiqueta = ArrParametros(1)
            EncriptStrconexion = CBool(ArrParametros(2))
            c_seed = ArrParametros(2)
        Else
            Exit Sub
        End If
    Else
        If UBound(ArrParametros) > 0 Then
            If IsNumeric(ArrParametros(0)) Then
                NroProcesoBatch = ArrParametros(0)
                Etiqueta = ArrParametros(1)
            Else
                Exit Sub
            End If
        Else
            If IsNumeric(strCmdLine) Then
                NroProcesoBatch = strCmdLine
            Else
                Exit Sub
            End If
        End If
    End If
    
    ' carga las configuraciones basicas, formato de fecha, string de conexion,
    ' tipo de BD y ubicacion del archivo de log
    Call CargarConfiguracionesBasicas

    On Error GoTo ME_Main
    Nombre_Arch = PathFLog & "Rep_listado_diario_asist" & "-" & NroProcesoBatch & ".log"
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Nombre_Arch, True)
    
    
    ' Obtengo el Process ID
    PID = GetCurrentProcessId
    Flog.writeline "-------------------------------------------------"
    Flog.writeline "Version                  : " & Version
    Flog.writeline "Fecha Ultima Modificacion: " & FechaVersion
    Flog.writeline "PID                      : " & PID
    Flog.writeline "-------------------------------------------------"
    Flog.writeline
    
    On Error Resume Next
    'Abro la conexion
    OpenConnection strconexion, objConn
    If Err.Number <> 0 Then
        Flog.writeline "Problemas en la conexion"
        Exit Sub
    End If
    OpenConnection strconexion, objconnProgreso
    If Err.Number <> 0 Then
        Flog.writeline "Problemas en la conexion"
        Exit Sub
    End If
    On Error GoTo 0
    
    On Error GoTo ME_Main
    
    
    'Cambio el estado del proceso a Procesando
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcprogreso = 0 ,bprcestado = 'Procesando', bprcpid = " & PID & " WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    'Flog.Writeline "Pone el estado en procesando"
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE btprcnro = 301 AND bpronro =" & NroProcesoBatch
    OpenRecordset StrSql, rs_batch_proceso
    
    TiempoInicialProceso = GetTickCount
    bprcparam = rs_batch_proceso!bprcparam
    If Not rs_batch_proceso.EOF Then
        bprcparam = rs_batch_proceso!bprcparam
        usuario = rs_batch_proceso!iduser
        rs_batch_proceso.Close
        Set rs_batch_proceso = Nothing
        NroProcesoBatch = NroProcesoBatch
        Call Listado_diario_asistencia(NroProcesoBatch, bprcparam)
    Else
        Flog.writeline "No encontró el proceso"
    End If
    
    TiempoFinalProceso = GetTickCount
    Flog.writeline "Tiempo del proceso (milisegundos): " & (TiempoFinalProceso - TiempoInicialProceso)
    
    If Not HuboError Then
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado' WHERE bpronro = " & NroProcesoBatch
    Else
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Error' WHERE bpronro = " & NroProcesoBatch
    End If
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
Fin:
    Flog.Close
    If objconnProgreso.State = adStateOpen Then objconnProgreso.Close
Exit Sub

ME_Main:
        
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
        StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcprogreso = 0, bprcestado = 'Error General', bprcpid = " & PID & " WHERE bpronro = " & NroProcesoBatch
        objconnProgreso.Execute StrSql, , adExecuteNoRecords
    MyCommitTrans
End Sub







Public Function nombre_estructura(ByVal Estrnro)
'DEVUELVE NOMBRE DE LA ESTDUCTURA

Dim rs_estr As New ADODB.Recordset

StrSql = "SELECT estrnro,estrdabr FROM estructura "
StrSql = StrSql & " WHERE "
StrSql = StrSql & " estrnro = " & Estrnro
OpenRecordset StrSql, rs_estr

If Not rs_estr.EOF Then
    If IsNull(rs_estr!estrdabr) Then
        nombre_estructura = ""
    Else
        nombre_estructura = rs_estr!estrdabr
    End If
    
Else
    nombre_estructura = ""
End If

rs_estr.Close

End Function

Public Function tipodehora_p1(ByVal Ternro, ByVal th, ByVal Fecha)
'DEVUELVE total de cantidad empleados

Dim rs_tipodehoras As New ADODB.Recordset
Dim a
a = 0
Ternro = Split(Ternro, ",")
StrSql = "SELECT COUNT(adcanthoras) adcanthoras FROM gti_acumdiario"


StrSql = StrSql & " WHERE (Ternro =" & Ternro(0)
For a = 1 To UBound(Ternro)
    StrSql = StrSql & " OR Ternro = " & Ternro(a)
Next

StrSql = StrSql & ") AND adfecha = " & ConvFecha(Fecha)
StrSql = StrSql & " AND thnro IN (" & th & ")"
 
OpenRecordset StrSql, rs_tipodehoras

If Not rs_tipodehoras.EOF Then
    If IsNull(rs_tipodehoras!adcanthoras) Then
        tipodehora_p1 = 0
    Else
        tipodehora_p1 = Replace(rs_tipodehoras!adcanthoras, ",", ".") & "@ "
        'Flog.writeline tipodehora
    End If
Else
    tipodehora_p1 = 0
End If

rs_tipodehoras.Close

End Function

Public Function tipodehora(ByVal Ternro, ByVal th, ByVal Fecha)
'DEVUELVE total de cantidad de horas del AD

Dim rs_tipodehoras As New ADODB.Recordset
Dim aux
StrSql = "SELECT SUM(adcanthoras) adcanthoras FROM gti_acumdiario"
'StrSql = "SELECT horas FROM gti_acumdiario"
StrSql = StrSql & " WHERE Ternro = " & Ternro
StrSql = StrSql & " AND adfecha = " & ConvFecha(Fecha)
StrSql = StrSql & " AND thnro IN (" & th & ")"
 
OpenRecordset StrSql, rs_tipodehoras

If Not rs_tipodehoras.EOF Then
    If IsNull(rs_tipodehoras!adcanthoras) Then
    'If IsNull(rs_tipodehoras!Horas) Then
        tipodehora = 0
    Else
        'tipodehora = Replace(rs_tipodehoras!adcanthoras, ",", ".")
        aux = 0
        tipodehora = CHoras(rs_tipodehoras!adcanthoras, 0)
        'Convierto hs en decimal
        'aux = Replace(rs_tipodehoras!Horas, ":", "")
        'tipodehora = aux / 1000
       'Flog.writeline tipodehora
    End If
Else
    tipodehora = 0
End If

rs_tipodehoras.Close

End Function
Public Function horcumplido(ByVal Ternro, ByVal th, ByVal Fecha)
'DEVUELVE NOMBRE DEL SUPERVISOR
Dim rs_horcumplido As New ADODB.Recordset
Dim empreporta

'StrSql = "SELECT empleado.empreporta FROM gti_horcumplido"
'StrSql = StrSql & " INNER JOIN empleado on empleado.ternro = gti_horcumplido.ternro "
StrSql = "SELECT ternro FROM gti_horcumplido"
StrSql = StrSql & " WHERE gti_horcumplido.Ternro = " & Ternro
StrSql = StrSql & " AND gti_horcumplido.thnro in (" & th & ")"
StrSql = StrSql & " AND (gti_horcumplido.hordesde = " & ConvFecha(Fecha) & " OR gti_horcumplido.horhasta = " & ConvFecha(Fecha) & ")"
OpenRecordset StrSql, rs_horcumplido

If Not rs_horcumplido.EOF Then
    Ternro = rs_horcumplido!Ternro
Else
    Ternro = ""
    horcumplido = ""
End If
rs_horcumplido.Close

'Busca el nombre del supervisor
If Ternro <> "" Then
    StrSql = "SELECT empreporta "
    StrSql = StrSql & "FROM empleado WHERE ternro = " & Ternro
    OpenRecordset StrSql, rs_horcumplido
    If Not rs_horcumplido.EOF Then
        empreporta = rs_horcumplido!empreporta
    Else
        empreporta = ""
    End If
    rs_horcumplido.Close
End If



If empreporta <> "" Then
    StrSql = "SELECT empleg,ternro,terape,ternom,terape2,ternom2 "
    StrSql = StrSql & " FROM empleado WHERE ternro = " & empreporta
    OpenRecordset StrSql, rs_horcumplido
    If Not rs_horcumplido.EOF Then
        horcumplido = rs_horcumplido!empleg & "@" & rs_horcumplido!ternom & "@" & rs_horcumplido!terape
    Else
        horcumplido = ""
    End If
    rs_horcumplido.Close
End If



End Function

Public Function lic_anual_p1(ByVal Ternro, ByVal tdnro, ByVal Fecha)
'DEVUELVE total de licencias tomadas por X empleados + descripción
Dim rs_lic_anual As New ADODB.Recordset
Dim c
Dim a
a = 0
c = 0
Ternro = Split(Ternro, ",")

StrSql = "SELECT elfechadesde,elfechahasta,tddesc,tdinteger4 "
StrSql = StrSql & " From emp_lic"
StrSql = StrSql & " INNER JOIN tipdia ON tipdia.tdnro = emp_lic.tdnro"
StrSql = StrSql & " WHERE (Empleado = " & Ternro(0)
For a = 1 To UBound(Ternro)
    StrSql = StrSql & " OR Empleado = " & Ternro(a)
Next

StrSql = StrSql & ") AND elfechadesde <= " & ConvFecha(Fecha) & " and elfechahasta >= " & ConvFecha(Fecha)
StrSql = StrSql & " AND emp_lic.tdnro IN (" & tdnro & ")"
OpenRecordset StrSql, rs_lic_anual
    
If Not rs_lic_anual.EOF Then
    Do While Not rs_lic_anual.EOF
        c = c + 1
        lic_anual_p1 = c & "@" & rs_lic_anual!tddesc
        rs_lic_anual.MoveNext
    Loop
Else
    lic_anual_p1 = ""
End If
rs_lic_anual.Close

End Function
Public Function lic_anual(ByVal Ternro, ByVal tdnro, ByVal all, ByVal Fecha)
'DEVUELVE total de cantidad de horas del AD

Dim rs_lic_anual As New ADODB.Recordset
Dim primerdiadelanio
Dim elfechadesde
Dim elfechahasta
Dim tddesc
Dim total_diastomados
Dim topeanual
Dim tdnro_aux
Dim c
total_diastomados = 0
tdnro_aux = Split(tdnro, ",")


For c = 0 To UBound(tdnro_aux)
    If all = "VAC" Then
        'Si la licencia es de vacaciones.
        StrSql = "SELECT elfechadesde,elfechahasta,tddesc,tdinteger4,vacacion.vacdesc descripcion"
        StrSql = StrSql & " From emp_lic"
        StrSql = StrSql & " INNER JOIN tipdia ON tipdia.tdnro = emp_lic.tdnro"
        StrSql = StrSql & " LEFT JOIN lic_vacacion ON lic_vacacion.emp_licnro = emp_lic.emp_licnro"
        StrSql = StrSql & " LEFT JOIN vacacion  ON vacacion.vacnro = lic_vacacion.vacnro"
        StrSql = StrSql & " WHERE Empleado = " & Ternro
        StrSql = StrSql & " AND elfechadesde <= " & ConvFecha(Fecha) & " and elfechahasta >= " & ConvFecha(Fecha)
    Else
        StrSql = "SELECT elfechadesde,elfechahasta,tddesc descripcion,tdinteger4 "
        StrSql = StrSql & " From emp_lic"
        StrSql = StrSql & " INNER JOIN tipdia ON tipdia.tdnro = emp_lic.tdnro"
        StrSql = StrSql & " WHERE Empleado = " & Ternro
        StrSql = StrSql & " AND elfechadesde <= " & ConvFecha(Fecha) & " and elfechahasta >= " & ConvFecha(Fecha)
    End If
    'Si all = 1 busca todas las licencias
    If all <> 1 Then
        StrSql = StrSql & " AND tipdia.tdnro IN (" & tdnro_aux(c) & ")"
    End If
    OpenRecordset StrSql, rs_lic_anual
            
    If Not rs_lic_anual.EOF Then
        elfechadesde = rs_lic_anual!elfechadesde
        elfechahasta = rs_lic_anual!elfechahasta
        
        'MMM - 08/05/2012
        topeanual = IIf(EsNulo(rs_lic_anual!tdinteger4), 0, rs_lic_anual!tdinteger4)
        
        'tddesc = rs_lic_anual!tddesc
        tddesc = rs_lic_anual!descripcion
        Exit For
    Else
        elfechadesde = ""
        elfechahasta = ""
        topeanual = ""
        tddesc = ""
    End If
   
Next

rs_lic_anual.Close

'Guarda 1er dia del año
primerdiadelanio = "01/01/" & Year(Fecha)


'Cuento cantidad de dias tomados
If elfechahasta <> "" Then
    StrSql = " Select SUM(elcantdias) elcantdias"
    StrSql = StrSql & " FROM emp_lic"
    StrSql = StrSql & " WHERE Empleado = " & Ternro
    StrSql = StrSql & " AND elfechadesde >= " & ConvFecha(primerdiadelanio) & " AND elfechahasta >= " & ConvFecha(elfechahasta)
    OpenRecordset StrSql, rs_lic_anual
    'Agregado la condicion del Len() por Carmen Quintero para que considere el caso, cuando la suma de cantidad de dias venga vacia
    If Not rs_lic_anual.EOF And Len(rs_lic_anual!elcantdias) > 0 Then
        total_diastomados = rs_lic_anual!elcantdias
    Else
        total_diastomados = 0
    End If
    
rs_lic_anual.Close
End If

If elfechadesde <> "" Then
   lic_anual = tddesc & " - " & elfechahasta & "@" & DateDiff("d", elfechadesde, Fecha) & "@" & topeanual - total_diastomados & "@" & total_diastomados
Else
    lic_anual = ""
End If

End Function

Public Function tipodehora_acum(ByVal Ternro, ByVal th, ByVal Fecha, ByVal tipo)
'DEVUELVE cantidad total de horas a una fecha

Dim rs_tipodehoras As New ADODB.Recordset


StrSql = "SELECT SUM(adcanthoras) adcanthoras FROM gti_acumdiario"
StrSql = StrSql & " WHERE Ternro = " & Ternro
StrSql = StrSql & " AND adfecha <= " & ConvFecha(Fecha)
StrSql = StrSql & " AND adfecha >= " & ConvFecha("01/01/" & Year(Fecha))
StrSql = StrSql & " AND thnro IN (" & th & ")"
 
OpenRecordset StrSql, rs_tipodehoras

If Not rs_tipodehoras.EOF Then
    If IsNull(rs_tipodehoras!adcanthoras) Then
        tipodehora_acum = 0
    Else
        If tipo = "TH" Then
            tipodehora_acum = CHoras(rs_tipodehoras!adcanthoras, 0)
            Flog.writeline "convierto a horas"
        Else
            tipodehora_acum = Replace(rs_tipodehoras!adcanthoras, ",", ".")
        End If
    End If
Else
    tipodehora_acum = 0
End If

rs_tipodehoras.Close

End Function



Public Function total_columna(ByVal repnro, ByVal coldesc, ByVal nrocol)
'DEVUELVE cantidad total de empleados para un tipo de columna + Observaciones
Dim rs_total As New ADODB.Recordset
Dim total_columna_aux As String

StrSql = "SELECT count(ternro) total FROM rep_list_asist_dia_p2"
StrSql = StrSql & " WHERE  repnro = " & repnro
StrSql = StrSql & " AND col" & nrocol & "_desc like '%" & coldesc & "@%'"
OpenRecordset StrSql, rs_total

If Not rs_total.EOF Then
  total_columna_aux = rs_total!total
Else
    total_columna_aux = "0"
End If

rs_total.Close

If total_columna <> "0" Then
    StrSql = "SELECT obs  FROM rep_list_asist_dia_p2"
    StrSql = StrSql & " WHERE  repnro = " & repnro
    StrSql = StrSql & " AND col" & nrocol & "_desc like '%" & coldesc & "@%'"
    OpenRecordset StrSql, rs_total

    If Not rs_total.EOF Then
        total_columna_aux = total_columna_aux & "@" & rs_total!obs
    Else
        total_columna_aux = total_columna_aux & "@"
    End If

End If
total_columna = total_columna_aux
rs_total.Close

End Function

Public Function hayregistracion(ByVal Ternro, ByVal Fecha)
    Dim rs_tienereg As New ADODB.Recordset

    'Si no tiene registraciones grabo X
    StrSql = "SELECT * FROM gti_registracion "
    StrSql = StrSql & " WHERE regfecha = " & ConvFecha(Fecha)
    StrSql = StrSql & " AND ternro = " & Ternro
    OpenRecordset StrSql, rs_tienereg
    If rs_tienereg.EOF Then
        hayregistracion = "X"
        Flog.writeline "Sin regitracion a horas"
    Else
        hayregistracion = ""
    End If
    rs_tienereg.Close
End Function

Public Function CHoras(ByVal Cantidad As Single, ByVal Dur As Single) As String
' ---------------------------------------------------------------------------------------------
' Descripcion: Retorna un string con la cantidad de hs y minutos a partir de un valor decimal
' Autor      : FGZ
' Fecha      : 09/11/2009
' Ultima Mod.:
' ---------------------------------------------------------------------------------------------
Dim Minutos As Single
Dim Horas As Single
    If Dur = 0 Then
        Dur = 60
    End If
    
    Cantidad = Cantidad * Dur
    Horas = Int(Cantidad / Dur)
    Minutos = Cantidad Mod Dur
    CHoras = "'" & Format(Horas, "#####00") & ":" & Format(Minutos, "00") & "'"
End Function



Public Function novedad_p1(ByVal Ternro, ByVal gtnovnro, ByVal Fecha)

Dim rs_novedad As New ADODB.Recordset
Dim gtnovnro_aux
Dim a
a = 0
gtnovnro_aux = 0
Ternro = Split(Ternro, ",")

'SI TIENE NOVEDAD ASOCIADA
    StrSql = "SELECT gti_novedad.gtnovnro,gti_tiponovedad.gtnovdesabr,gti_novedad.gnovdesext"
    StrSql = StrSql & " FROM gti_justificacion"
    StrSql = StrSql & " INNER JOIN  gti_novedad ON gti_justificacion.juscodext= gti_novedad.gnovnro"
    StrSql = StrSql & " INNER JOIN gti_tiponovedad ON gti_tiponovedad.gtnovnro = gti_novedad.gtnovnro"
    
    StrSql = StrSql & " WHERE (gti_justificacion.Ternro = " & Ternro(0)
    For a = 1 To UBound(Ternro)
        StrSql = StrSql & " OR gti_justificacion.Ternro = " & Ternro(a)
    Next
    
    StrSql = StrSql & ") AND gti_justificacion.jusdesde <= " & ConvFecha(Fecha) & " and gti_justificacion.jushasta >=" & ConvFecha(Fecha)
    StrSql = StrSql & " AND gti_novedad.gtnovnro IN(" & gtnovnro & ")"
    StrSql = StrSql & " AND gti_justificacion.jussigla = 'NOV'"
    OpenRecordset StrSql, rs_novedad
    
    If Not rs_novedad.EOF Then
        
        Do While Not rs_novedad.EOF
            'cuento cantidad de registros
            gtnovnro_aux = gtnovnro_aux + 1
            novedad_p1 = gtnovnro_aux & "@" & Trim(rs_novedad!gtnovdesabr) & " - " & Trim(rs_novedad!gnovdesext)
            rs_novedad.MoveNext
        Loop
        
    Else
        novedad_p1 = ""
    End If
    
    rs_novedad.Close




End Function
Public Function novedad(ByVal Ternro, ByVal gtnovnro, ByVal Fecha)
'
Dim rs_novedad As New ADODB.Recordset

Dim gtnovnro_aux

gtnovnro_aux = 0


'SI TIENE NOVEDAD ASOCIADA
StrSql = "SELECT gti_novedad.gtnovnro,gti_tiponovedad.gtnovdesabr,gti_novedad.gnovdesext"
StrSql = StrSql & " FROM gti_justificacion"
StrSql = StrSql & " INNER JOIN  gti_novedad ON gti_justificacion.juscodext= gti_novedad.gnovnro"
StrSql = StrSql & " INNER JOIN gti_tiponovedad ON gti_tiponovedad.gtnovnro = gti_novedad.gtnovnro"
StrSql = StrSql & " WHERE gti_justificacion.Ternro = " & Ternro
StrSql = StrSql & " AND gti_justificacion.jusdesde <= " & ConvFecha(Fecha) & " and gti_justificacion.jushasta >=" & ConvFecha(Fecha)
StrSql = StrSql & " AND gti_novedad.gtnovnro IN(" & gtnovnro & ")"
StrSql = StrSql & " AND gti_justificacion.jussigla = 'NOV'"
OpenRecordset StrSql, rs_novedad

If Not rs_novedad.EOF Then
    novedad = Trim(rs_novedad!gtnovdesabr) & " - " & Trim(rs_novedad!gnovdesext)
Else
    novedad = ""
End If
    
rs_novedad.Close


'Concateno thnro
If novedad <> "" Then
    StrSql = "SELECT thnro FROM gti_tiponovedad "
    StrSql = StrSql & " WHERE gtnovnro IN ( " & gtnovnro & ")"
    OpenRecordset StrSql, rs_novedad
    If Not rs_novedad.EOF Then
        novedad = novedad & "@" & rs_novedad!thnro
    End If
End If

End Function

Public Function estructura_actual(ByVal Ternro, ByVal empleg, ByVal Fecha)
'DEVUELVE NOMBRE DE LA ESTRUCTURA
Dim rs_estr As New ADODB.Recordset

StrSql = "SELECT his_estructura.estrnro "
StrSql = StrSql & " FROM his_estructura "
StrSql = StrSql & " INNER JOIN empleado ON his_estructura.ternro = empleado.ternro "
StrSql = StrSql & " WHERE his_estructura.tenro =  " & Ternro
StrSql = StrSql & " AND  ( his_estructura.htethasta IS NULL "
StrSql = StrSql & " OR his_estructura.htethasta >= " & ConvFecha(Fecha) & ") "
StrSql = StrSql & " AND his_estructura.htetdesde <= " & ConvFecha(Fecha)
StrSql = StrSql & " AND empleado.empleg = " & empleg
OpenRecordset StrSql, rs_estr

If Not rs_estr.EOF Then
    If IsNull(rs_estr!Estrnro) Then
        estructura_actual = 0
    Else
        estructura_actual = rs_estr!Estrnro
    End If
    
Else
    estructura_actual = 0
End If

rs_estr.Close

End Function

Public Sub Listado_diario_asistencia(ByVal bpronro As Long, ByVal parametros As String)
' --------------------------------------------------------------------------------------------
' Descripcion: Listado Diario de Asistencia
' Autor      : Gonzalez Nicolás
' Fecha      : 23/05/2011
' Descripción: Reporte - Parte1 - Parte 2 - Parte 3
' --------------------------------------------------------------------------------------------

Dim pos1 As Integer
'Dim pos2 As Integer
'Dim Nroliq As Integer
Dim arr_lista

'Progreso
Dim CEmpleadosAProc As Long
Dim IncPorc As Double
Dim Progreso As Double

'Vars. grales.
Dim nombre As String
Dim Apellido As String

Dim StrSql_ins
'Parámetros de Batch_proceso
Dim Emp_desde
Dim Emp_hasta
Dim Estado
Dim t_nivel1
Dim t_nivel2
Dim t_nivel3
Dim Fecha
Dim Empleados
Dim Orden
Dim Ordenado
Dim t_estruc1
Dim t_estruc2
Dim t_estruc3
'------------

Dim rs As New ADODB.Recordset
Dim rs_supervisor As New ADODB.Recordset
Dim StrSql2
Dim Select_empleados
'-------------------

'Tipo estructuras Parte 2
Dim col_nombre1
Dim testrnro1
Dim col_nombre2
Dim testrnro2
Dim col_nombre3
Dim testrnro3
Dim col_nombre4
Dim testrnro4
Dim col_nombre5
Dim testrnro5

'Tipo estructuras Parte 3
Dim col_nombre6
Dim testrnro6
Dim col_nombre7
Dim testrnro7
Dim col_nombre8
Dim testrnro8
Dim col_nombre9
Dim testrnro9
Dim col_nombre10
Dim testrnro10


Dim historico
'Valor default (En caso de que no este configurado) - Para eliminar historico
historico = 10

Dim tipo As String
tipo = ""

'Columnas de 1 a 20
'Setea valor x cantidad de columnas máximas
Dim columna(19, 19, 19)
Dim col(19)
Dim confval2(19)
Dim confval2_p1(19)
Dim columna_p1(19, 19, 19)


'N° Estruc y Descripción
Dim estrnro1
Dim estrnro2
Dim estrnro3
Dim estrnro4
Dim estrnro5
Dim estrnro6
Dim estrnro7
Dim estrnro8
Dim estrnro9
Dim estrnro10

Dim estrnro1_desc
Dim estrnro2_desc
Dim estrnro3_desc
Dim estrnro4_desc
Dim estrnro5_desc
Dim estrnro6_desc
Dim estrnro7_desc
Dim estrnro8_desc
Dim estrnro9_desc
Dim estrnro10_desc

estrnro1 = 0
estrnro2 = 0
estrnro3 = 0
estrnro4 = 0
estrnro5 = 0
estrnro6 = 0
estrnro7 = 0
estrnro8 = 0
estrnro9 = 0
estrnro10 = 0

estrnro1_desc = ""
estrnro2_desc = ""
estrnro3_desc = ""
estrnro4_desc = ""
estrnro5_desc = ""
estrnro6_desc = ""
estrnro7_desc = ""
estrnro8_desc = ""
estrnro9_desc = ""
estrnro10_desc = ""

'Seteo todos en vacio
col_nombre1 = ""
testrnro1 = 0
col_nombre2 = ""
testrnro2 = 0
col_nombre3 = ""
testrnro3 = 0
col_nombre4 = ""
testrnro4 = 0
col_nombre5 = ""
testrnro5 = 0
col_nombre6 = ""
testrnro6 = 0
col_nombre7 = ""
testrnro7 = 0
col_nombre8 = ""
testrnro8 = 0
col_nombre9 = ""
testrnro9 = 0
col_nombre10 = ""
testrnro10 = 0


Dim observaciones
observaciones = ""

Dim nro_col As Integer
nro_col = 0

Dim acumulado
Dim acumulado_total
acumulado_total = "NULL"
acumulado = 0
Dim repnro
Dim Texto_aux
Texto_aux = ""

'Para licencia anual
Dim licencia_anual
Dim c

'FORMATEO VALORES EN NULL
For c = 0 To 19
    col(c) = "NULL"
Next

Dim novedad_aux
'¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
'REPORTE PARTE  2
'¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
Dim total_x_col
Dim total_x_col2
total_x_col = 0
Dim descripcion_p1

Dim cant_empleados As Integer
cant_empleados = 0
Dim Ternro_empleados
Ternro_empleados = "0"

'¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
'REPORTE PARTE  3
'¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
Dim datos_supervisor
Dim hs_ausencia
hs_ausencia = ""
datos_supervisor = ""


Dim hayreg
hayreg = ""
'----------------------------------------------------------------------
'Levanto cada parametro por separado, el separador es "@"
'----------------------------------------------------------------------
'Ejemplo: 1@2@1000000000@29@360@29@514@7@1299@10/05/2011@-1@Emp@Asc

Flog.writeline "Levantando parametros " & parametros


If Not IsNull(parametros) Then
    
    If Len(parametros) >= 1 Then
     
        arr_lista = Split(parametros, "@", -1, 1)
        '-----
        Emp_desde = arr_lista(1)
        Emp_hasta = arr_lista(2)
        Fecha = arr_lista(9)
        '-1 | 0 | 1
        If arr_lista(0) = "-1" Then
            'Activo
            Estado = "(empleado.ternro in (SELECT empleado from fases WHERE fases.empleado = empleado.ternro AND (fases.altfec <=" & ConvFecha(Fecha) & " AND (fases.bajfec is null or fases.bajfec>=" & ConvFecha(Fecha) & ")) )) AND (empleg >= " & Emp_desde & ") AND (empleg <= " & Emp_hasta & ") "
        ElseIf arr_lista(0) = "0" Then
            'Inactivo
            Estado = "(empleado.ternro not in (SELECT empleado from fases WHERE fases.empleado = empleado.ternro AND (fases.altfec <=" & ConvFecha(Fecha) & " AND (fases.bajfec is null or fases.bajfec>=" & ConvFecha(Fecha) & ")) )) AND (empleg >= " & Emp_desde & ") AND (empleg <= " & Emp_hasta & ") "
        Else
            'Ambos
            Estado = "(0 = 0) And (empleg >= " & Emp_desde & ") And (empleg <= " & Emp_hasta & ")"
        End If
        
        t_nivel1 = arr_lista(3)
        t_estruc1 = arr_lista(4)
        t_nivel2 = arr_lista(5)
        t_estruc2 = arr_lista(6)
        t_nivel3 = arr_lista(7)
        t_estruc3 = arr_lista(8)
        
        
        
        'Todos = 0 | Propios = -1| Agencia -2
        Empleados = arr_lista(10)
        
        Orden = arr_lista(11)
        Ordenado = arr_lista(12)
        
        If Orden = "Emp" Then
            Orden = "Empleg"
        Else
            Orden = "Terape"
        End If
    End If
Else
    Flog.writeline "Error en parámetros"
End If

Flog.writeline "Parámetros levantados correctamente"
Flog.writeline ""

'--------------------------------------------------------------------
'TRAE DATOS FIJOS CONFIGURABLES CONFREP
'--------------------------------------------------------------------
StrSql = "SELECT * FROM  confrep WHERE repnro = 348 "
StrSql = StrSql & " ORDER BY confnrocol "
OpenRecordset StrSql, rs

If rs.EOF Then
    Flog.writeline "No se encontró la configuración del Reporte"
    Exit Sub
Else
    Do While Not rs.EOF
        '_______________________________
        'ESTRUCTURAS PARTE 2
        '_______________________________
        Select Case Trim(rs!conftipo) & Trim(rs!confval2)
            Case "TE1P2":
                col_nombre1 = rs!confetiq
                testrnro1 = rs!confval
            Case "TE2P2":
                col_nombre2 = rs!confetiq
                testrnro2 = rs!confval
            Case "TE3P2":
                col_nombre3 = rs!confetiq
                testrnro3 = rs!confval
            Case "TE4P2":
                col_nombre4 = rs!confetiq
                testrnro4 = rs!confval
            Case "TE5P2":
                col_nombre5 = rs!confetiq
                testrnro5 = rs!confval
         'End Select
         
         
                 '_______________________________
        'ESTRUCTURAS PARTE 3
        '_______________________________
        'Select Case Trim(rs!conftipo) & Trim(rs!confval2)
            Case "TE1P3":
                col_nombre6 = rs!confetiq
                testrnro6 = rs!confval
            Case "TE2P3":
                col_nombre7 = rs!confetiq
                testrnro7 = rs!confval
            Case "TE3P3":
                col_nombre8 = rs!confetiq
                testrnro8 = rs!confval
            Case "TE4P3":
                col_nombre9 = rs!confetiq
                testrnro9 = rs!confval
            Case "TE5P3":
                col_nombre10 = rs!confetiq
                testrnro10 = rs!confval
            
            'Histórico
            Case "HIS"
                historico = rs!confval
            Case "THHR"
                hs_ausencia = rs!confval
         End Select
         
         '_______________________________
         'COLUMNAS PARTE 2
         '_______________________________
         'Guardo el código de cada tipo de hora
         Select Case rs!confnrocol
            Case "1"
                If columna(0, 0, 0) = "" Then
                    columna(0, 0, 0) = Trim(rs!confval)
                    columna(0, 1, 0) = Trim(rs!conftipo)
                    columna(0, 1, 1) = Trim(rs!confetiq)
                    confval2(0) = Trim(rs!confval2)
                ElseIf columna(0, 1, 0) <> Trim(rs!conftipo) Then
                    Flog.writeline "Error de tipo de datos en columna N°: 1 " & Trim(rs!conftipo) & " No coincide"
                Else
                    columna(0, 0, 0) = columna(0, 0, 0) & "," & rs!confval
                End If
                
            Case "2"
                If columna(1, 0, 0) = "" Then
                    columna(1, 0, 0) = Trim(rs!confval)
                    columna(1, 1, 0) = Trim(rs!conftipo)
                    columna(1, 1, 1) = Trim(rs!confetiq)
                    confval2(1) = Trim(rs!confval2)
                    
                ElseIf columna(1, 1, 0) <> Trim(rs!conftipo) Then
                    Flog.writeline "Error de tipo de datos en columna N°: 2 " & Trim(rs!conftipo) & " No coincide"
                Else
                    columna(1, 0, 0) = columna(1, 0, 0) & "," & rs!confval
                End If
                
            Case "3"
                If columna(2, 0, 0) = "" Then
                    columna(2, 0, 0) = Trim(rs!confval)
                    columna(2, 1, 0) = Trim(rs!conftipo)
                    columna(2, 1, 1) = Trim(rs!confetiq)
                    confval2(2) = Trim(rs!confval2)
                ElseIf columna(2, 1, 0) <> Trim(rs!conftipo) Then
                    Flog.writeline "Error de tipo de datos en columna N°: 3 " & Trim(rs!conftipo) & " No coincide"
                Else
                    columna(2, 0, 0) = columna(2, 0, 0) & "," & rs!confval
                End If
                
            Case "4"
                If columna(3, 0, 0) = "" Then
                    columna(3, 0, 0) = Trim(rs!confval)
                    columna(3, 1, 0) = Trim(rs!conftipo)
                    columna(3, 1, 1) = Trim(rs!confetiq)
                    confval2(3) = Trim(rs!confval2)
                ElseIf columna(3, 1, 0) <> Trim(rs!conftipo) Then
                    Flog.writeline "Error de tipo de datos en columna N°: 4 " & Trim(rs!conftipo) & " No coincide"
                Else
                    columna(3, 0, 0) = columna(4, 0, 0) & "," & rs!confval
                End If
                
            Case "5"
                If columna(4, 0, 0) = "" Then
                    columna(4, 0, 0) = Trim(rs!confval)
                    columna(4, 1, 0) = Trim(rs!conftipo)
                    columna(4, 1, 1) = Trim(rs!confetiq)
                    confval2(4) = Trim(rs!confval2)
                ElseIf columna(4, 1, 0) <> Trim(rs!conftipo) Then
                    Flog.writeline "Error de tipo de datos en columna N°: 5 " & Trim(rs!conftipo) & " No coincide"
                Else
                    columna(4, 0, 0) = columna(4, 0, 0) & "," & rs!confval
                End If
            Case "6"
                If columna(5, 0, 0) = "" Then
                    columna(5, 0, 0) = Trim(rs!confval)
                    columna(5, 1, 0) = Trim(rs!conftipo)
                    columna(5, 1, 1) = Trim(rs!confetiq)
                    confval2(5) = Trim(rs!confval2)
                ElseIf columna(5, 1, 0) <> Trim(rs!conftipo) Then
                    Flog.writeline "Error de tipo de datos en columna N°: 6 " & Trim(rs!conftipo) & " No coincide"
                Else
                    columna(5, 0, 0) = columna(5, 0, 0) & "," & rs!confval
                End If
            Case "7"
                If columna(6, 0, 0) = "" Then
                    columna(6, 0, 0) = Trim(rs!confval)
                    columna(6, 1, 0) = Trim(rs!conftipo)
                    columna(6, 1, 1) = Trim(rs!confetiq)
                    confval2(6) = Trim(rs!confval2)
                ElseIf columna(6, 1, 0) <> Trim(rs!conftipo) Then
                    Flog.writeline "Error de tipo de datos en columna N°: 7 " & Trim(rs!conftipo) & " No coincide"
                Else
                    columna(6, 0, 0) = columna(6, 0, 0) & "," & rs!confval
                End If
            Case "8"
                If columna(7, 0, 0) = "" Then
                    columna(7, 0, 0) = Trim(rs!confval)
                    columna(7, 1, 0) = Trim(rs!conftipo)
                    columna(7, 1, 1) = Trim(rs!confetiq)
                    confval2(7) = Trim(rs!confval2)
                ElseIf columna(7, 1, 0) <> Trim(rs!conftipo) Then
                    Flog.writeline "Error de tipo de datos en columna N°: 8 " & Trim(rs!conftipo) & " No coincide"
                Else
                    columna(7, 0, 0) = columna(7, 0, 0) & "," & rs!confval
                End If
            Case "9"
                If columna(8, 0, 0) = "" Then
                    columna(8, 0, 0) = Trim(rs!confval)
                    columna(8, 1, 0) = Trim(rs!conftipo)
                    columna(8, 1, 1) = Trim(rs!confetiq)
                    confval2(8) = Trim(rs!confval2)
                ElseIf columna(8, 1, 0) <> Trim(rs!conftipo) Then
                    Flog.writeline "Error de tipo de datos en columna N°: 9 " & Trim(rs!conftipo) & " No coincide"
                Else
                    columna(8, 0, 0) = columna(8, 0, 0) & "," & rs!confval
                End If
            Case "10"
                If columna(9, 0, 0) = "" Then
                    columna(9, 0, 0) = Trim(rs!confval)
                    columna(9, 1, 0) = Trim(rs!conftipo)
                    columna(9, 1, 1) = Trim(rs!confetiq)
                    confval2(9) = Trim(rs!confval2)
                ElseIf columna(9, 1, 0) <> Trim(rs!conftipo) Then
                    Flog.writeline "Error de tipo de datos en columna N°: 10 " & Trim(rs!conftipo) & " No coincide"
                Else
                    columna(9, 0, 0) = columna(9, 0, 0) & "," & rs!confval
                End If
            Case "11"
                If columna(10, 0, 0) = "" Then
                    columna(10, 0, 0) = Trim(rs!confval)
                    columna(10, 1, 0) = Trim(rs!conftipo)
                    columna(10, 1, 1) = Trim(rs!confetiq)
                    confval2(10) = Trim(rs!confval2)
                ElseIf columna(10, 1, 0) <> Trim(rs!conftipo) Then
                    Flog.writeline "Error de tipo de datos en columna N°: 11 " & Trim(rs!conftipo) & " No coincide"
                Else
                    columna(10, 0, 0) = columna(10, 0, 0) & "," & rs!confval
                End If
            Case "12"
                If columna(11, 0, 0) = "" Then
                    columna(11, 0, 0) = Trim(rs!confval)
                    columna(11, 1, 0) = Trim(rs!conftipo)
                    columna(11, 1, 1) = Trim(rs!confetiq)
                    confval2(11) = Trim(rs!confval2)
                ElseIf columna(11, 1, 0) <> Trim(rs!conftipo) Then
                    Flog.writeline "Error de tipo de datos en columna N°: 12 " & Trim(rs!conftipo) & " No coincide"
                Else
                    columna(11, 0, 0) = columna(11, 0, 0) & "," & rs!confval
                End If
            Case "13"
                If columna(12, 0, 0) = "" Then
                    columna(12, 0, 0) = Trim(rs!confval)
                    columna(12, 1, 0) = Trim(rs!conftipo)
                    columna(12, 1, 1) = Trim(rs!confetiq)
                    confval2(12) = Trim(rs!confval2)
                ElseIf columna(12, 1, 0) <> Trim(rs!conftipo) Then
                    Flog.writeline "Error de tipo de datos en columna N°: 13 " & Trim(rs!conftipo) & " No coincide"
                Else
                    columna(12, 0, 0) = columna(12, 0, 0) & "," & rs!confval
                End If
            Case "14"
                If columna(13, 0, 0) = "" Then
                    columna(13, 0, 0) = Trim(rs!confval)
                    columna(13, 1, 0) = Trim(rs!conftipo)
                    columna(13, 1, 1) = Trim(rs!confetiq)
                    confval2(13) = Trim(rs!confval2)
                ElseIf columna(13, 1, 0) <> Trim(rs!conftipo) Then
                    Flog.writeline "Error de tipo de datos en columna N°: 14 " & Trim(rs!conftipo) & " No coincide"
                Else
                    columna(13, 0, 0) = columna(13, 0, 0) & "," & rs!confval
                End If
            Case "15"
                If columna(14, 0, 0) = "" Then
                    columna(14, 0, 0) = Trim(rs!confval)
                    columna(14, 1, 0) = Trim(rs!conftipo)
                    columna(14, 1, 1) = Trim(rs!confetiq)
                    confval2(14) = Trim(rs!confval2)
                ElseIf columna(14, 1, 0) <> Trim(rs!conftipo) Then
                    Flog.writeline "Error de tipo de datos en columna N°: 15 " & Trim(rs!conftipo) & " No coincide"
                Else
                    columna(14, 0, 0) = columna(14, 0, 0) & "," & rs!confval
                End If
            Case "16"
                If columna(15, 0, 0) = "" Then
                    columna(15, 0, 0) = Trim(rs!confval)
                    columna(15, 1, 0) = Trim(rs!conftipo)
                    columna(15, 1, 1) = Trim(rs!confetiq)
                    confval2(15) = Trim(rs!confval2)
                ElseIf columna(15, 1, 0) <> Trim(rs!conftipo) Then
                    Flog.writeline "Error de tipo de datos en columna N°: 16 " & Trim(rs!conftipo) & " No coincide"
                Else
                    columna(15, 0, 0) = columna(15, 0, 0) & "," & rs!confval
                End If
            Case "17"
                If columna(16, 0, 0) = "" Then
                    columna(16, 0, 0) = Trim(rs!confval)
                    columna(16, 1, 0) = Trim(rs!conftipo)
                    columna(16, 1, 1) = Trim(rs!confetiq)
                    confval2(16) = Trim(rs!confval2)
                ElseIf columna(16, 1, 0) <> Trim(rs!conftipo) Then
                    Flog.writeline "Error de tipo de datos en columna N°: 17 " & Trim(rs!conftipo) & " No coincide"
                Else
                    columna(16, 0, 0) = columna(16, 0, 0) & "," & rs!confval
                End If
            Case "18"
                If columna(17, 0, 0) = "" Then
                    columna(17, 0, 0) = Trim(rs!confval)
                    columna(17, 1, 0) = Trim(rs!conftipo)
                    columna(17, 1, 1) = Trim(rs!confetiq)
                    confval2(17) = Trim(rs!confval2)
                ElseIf columna(17, 1, 0) <> Trim(rs!conftipo) Then
                    Flog.writeline "Error de tipo de datos en columna N°: 18 " & Trim(rs!conftipo) & " No coincide"
                Else
                    columna(17, 0, 0) = columna(17, 0, 0) & "," & rs!confval
                End If
            Case "19"
                If columna(18, 0, 0) = "" Then
                    columna(18, 0, 0) = Trim(rs!confval)
                    columna(18, 1, 0) = Trim(rs!conftipo)
                    columna(18, 1, 1) = Trim(rs!confetiq)
                    confval2(18) = Trim(rs!confval2)
                ElseIf columna(18, 1, 0) <> Trim(rs!conftipo) Then
                    Flog.writeline "Error de tipo de datos en columna N°: 19 " & Trim(rs!conftipo) & " No coincide"
                Else
                    columna(18, 0, 0) = columna(18, 0, 0) & "," & rs!confval
                End If
            Case "20"
                If columna(19, 0, 0) = "" Then
                    columna(19, 0, 0) = Trim(rs!confval)
                    columna(19, 1, 0) = Trim(rs!conftipo)
                    columna(19, 1, 1) = Trim(rs!confetiq)
                    confval2(19) = Trim(rs!confval2)
                ElseIf columna(19, 1, 0) <> Trim(rs!conftipo) Then
                    Flog.writeline "Error de tipo de datos en columna N°: 20 " & Trim(rs!conftipo) & " No coincide"
                Else
                    columna(19, 0, 0) = columna(19, 0, 0) & "," & rs!confval
                End If
                
        '_______________________________
        'FILAS PARTE 1
        '_______________________________
            Case "21"
                If columna_p1(0, 0, 0) = "" Then
                    columna_p1(0, 0, 0) = Trim(rs!confval)
                    columna_p1(0, 1, 0) = Trim(rs!conftipo)
                    columna_p1(0, 1, 1) = Trim(rs!confetiq)
                    confval2_p1(0) = Trim(rs!confval2)
                ElseIf columna_p1(0, 1, 0) <> Trim(rs!conftipo) Then
                    Flog.writeline "Error de tipo de datos en columna N°: 21 " & Trim(rs!conftipo) & " No coincide"
                Else
                    columna_p1(0, 0, 0) = columna_p1(0, 0, 0) & "," & rs!confval
                End If
                
            Case "22"
                If columna_p1(1, 0, 0) = "" Then
                    columna_p1(1, 0, 0) = Trim(rs!confval)
                    columna_p1(1, 1, 0) = Trim(rs!conftipo)
                    columna_p1(1, 1, 1) = Trim(rs!confetiq)
                    confval2_p1(1) = Trim(rs!confval2)
                    
                ElseIf columna_p1(1, 1, 0) <> Trim(rs!conftipo) Then
                    Flog.writeline "Error de tipo de datos en columna N°: 22 " & Trim(rs!conftipo) & " No coincide"
                Else
                    columna_p1(1, 0, 0) = columna_p1(1, 0, 0) & "," & rs!confval
                End If
                
            Case "23"
                If columna_p1(2, 0, 0) = "" Then
                    columna_p1(2, 0, 0) = Trim(rs!confval)
                    columna_p1(2, 1, 0) = Trim(rs!conftipo)
                    columna_p1(2, 1, 1) = Trim(rs!confetiq)
                    confval2_p1(2) = Trim(rs!confval2)
                ElseIf columna_p1(2, 1, 0) <> Trim(rs!conftipo) Then
                    Flog.writeline "Error de tipo de datos en columna N°: 23 " & Trim(rs!conftipo) & " No coincide"
                Else
                    columna_p1(2, 0, 0) = columna_p1(2, 0, 0) & "," & rs!confval
                End If
                
            Case "24"
                If columna_p1(3, 0, 0) = "" Then
                    columna_p1(3, 0, 0) = Trim(rs!confval)
                    columna_p1(3, 1, 0) = Trim(rs!conftipo)
                    columna_p1(3, 1, 1) = Trim(rs!confetiq)
                    confval2_p1(3) = Trim(rs!confval2)
                ElseIf columna_p1(3, 1, 0) <> Trim(rs!conftipo) Then
                    Flog.writeline "Error de tipo de datos en columna N°: 24 " & Trim(rs!conftipo) & " No coincide"
                Else
                    columna_p1(3, 0, 0) = columna_p1(4, 0, 0) & "," & rs!confval
                End If
                
            Case "25"
                If columna_p1(4, 0, 0) = "" Then
                    columna_p1(4, 0, 0) = Trim(rs!confval)
                    columna_p1(4, 1, 0) = Trim(rs!conftipo)
                    columna_p1(4, 1, 1) = Trim(rs!confetiq)
                    confval2_p1(4) = Trim(rs!confval2)
                ElseIf columna_p1(4, 1, 0) <> Trim(rs!conftipo) Then
                    Flog.writeline "Error de tipo de datos en columna N°: 25 " & Trim(rs!conftipo) & " No coincide"
                Else
                    columna_p1(4, 0, 0) = columna_p1(4, 0, 0) & "," & rs!confval
                End If
            Case "26"
                If columna_p1(5, 0, 0) = "" Then
                    columna_p1(5, 0, 0) = Trim(rs!confval)
                    columna_p1(5, 1, 0) = Trim(rs!conftipo)
                    columna_p1(5, 1, 1) = Trim(rs!confetiq)
                    confval2_p1(5) = Trim(rs!confval2)
                ElseIf columna_p1(5, 1, 0) <> Trim(rs!conftipo) Then
                    Flog.writeline "Error de tipo de datos en columna N°: 26 " & Trim(rs!conftipo) & " No coincide"
                Else
                    columna_p1(5, 0, 0) = columna_p1(5, 0, 0) & "," & rs!confval
                End If
            Case "27"
                If columna_p1(6, 0, 0) = "" Then
                    columna_p1(6, 0, 0) = Trim(rs!confval)
                    columna_p1(6, 1, 0) = Trim(rs!conftipo)
                    columna_p1(6, 1, 1) = Trim(rs!confetiq)
                    confval2_p1(6) = Trim(rs!confval2)
                ElseIf columna_p1(6, 1, 0) <> Trim(rs!conftipo) Then
                    Flog.writeline "Error de tipo de datos en columna N°: 27 " & Trim(rs!conftipo) & " No coincide"
                Else
                    columna_p1(6, 0, 0) = columna_p1(6, 0, 0) & "," & rs!confval
                End If
            Case "28"
                If columna_p1(7, 0, 0) = "" Then
                    columna_p1(7, 0, 0) = Trim(rs!confval)
                    columna_p1(7, 1, 0) = Trim(rs!conftipo)
                    columna_p1(7, 1, 1) = Trim(rs!confetiq)
                    confval2_p1(7) = Trim(rs!confval2)
                ElseIf columna_p1(7, 1, 0) <> Trim(rs!conftipo) Then
                    Flog.writeline "Error de tipo de datos en columna N°: 28 " & Trim(rs!conftipo) & " No coincide"
                Else
                    columna_p1(7, 0, 0) = columna_p1(7, 0, 0) & "," & rs!confval
                End If
            Case "29"
                If columna_p1(8, 0, 0) = "" Then
                    columna_p1(8, 0, 0) = Trim(rs!confval)
                    columna_p1(8, 1, 0) = Trim(rs!conftipo)
                    columna_p1(8, 1, 1) = Trim(rs!confetiq)
                    confval2_p1(8) = Trim(rs!confval2)
                ElseIf columna_p1(8, 1, 0) <> Trim(rs!conftipo) Then
                    Flog.writeline "Error de tipo de datos en columna N°: 29 " & Trim(rs!conftipo) & " No coincide"
                Else
                    columna_p1(8, 0, 0) = columna_p1(8, 0, 0) & "," & rs!confval
                End If
            Case "30"
                If columna_p1(9, 0, 0) = "" Then
                    columna_p1(9, 0, 0) = Trim(rs!confval)
                    columna_p1(9, 1, 0) = Trim(rs!conftipo)
                    columna_p1(9, 1, 1) = Trim(rs!confetiq)
                    confval2_p1(9) = Trim(rs!confval2)
                ElseIf columna_p1(9, 1, 0) <> Trim(rs!conftipo) Then
                    Flog.writeline "Error de tipo de datos en columna N°: 30 " & Trim(rs!conftipo) & " No coincide"
                Else
                    columna_p1(9, 0, 0) = columna_p1(9, 0, 0) & "," & rs!confval
                End If
            Case "31"
                If columna_p1(10, 0, 0) = "" Then
                    columna_p1(10, 0, 0) = Trim(rs!confval)
                    columna_p1(10, 1, 0) = Trim(rs!conftipo)
                    columna_p1(10, 1, 1) = Trim(rs!confetiq)
                    confval2_p1(10) = Trim(rs!confval2)
                ElseIf columna_p1(10, 1, 0) <> Trim(rs!conftipo) Then
                    Flog.writeline "Error de tipo de datos en columna N°: 31 " & Trim(rs!conftipo) & " No coincide"
                Else
                    columna_p1(10, 0, 0) = columna_p1(10, 0, 0) & "," & rs!confval
                End If
            Case "32"
                If columna_p1(11, 0, 0) = "" Then
                    columna_p1(11, 0, 0) = Trim(rs!confval)
                    columna_p1(11, 1, 0) = Trim(rs!conftipo)
                    columna_p1(11, 1, 1) = Trim(rs!confetiq)
                    confval2_p1(11) = Trim(rs!confval2)
                ElseIf columna_p1(11, 1, 0) <> Trim(rs!conftipo) Then
                    Flog.writeline "Error de tipo de datos en columna N°: 32 " & Trim(rs!conftipo) & " No coincide"
                Else
                    columna_p1(11, 0, 0) = columna_p1(11, 0, 0) & "," & rs!confval
                End If
            Case "33"
                If columna_p1(12, 0, 0) = "" Then
                    columna_p1(12, 0, 0) = Trim(rs!confval)
                    columna_p1(12, 1, 0) = Trim(rs!conftipo)
                    columna_p1(12, 1, 1) = Trim(rs!confetiq)
                    confval2_p1(12) = Trim(rs!confval2)
                ElseIf columna_p1(12, 1, 0) <> Trim(rs!conftipo) Then
                    Flog.writeline "Error de tipo de datos en columna N°: 33 " & Trim(rs!conftipo) & " No coincide"
                Else
                    columna_p1(12, 0, 0) = columna_p1(12, 0, 0) & "," & rs!confval
                End If
            Case "34"
                If columna_p1(13, 0, 0) = "" Then
                    columna_p1(13, 0, 0) = Trim(rs!confval)
                    columna_p1(13, 1, 0) = Trim(rs!conftipo)
                    columna_p1(13, 1, 1) = Trim(rs!confetiq)
                    confval2_p1(13) = Trim(rs!confval2)
                ElseIf columna_p1(13, 1, 0) <> Trim(rs!conftipo) Then
                    Flog.writeline "Error de tipo de datos en columna N°: 34 " & Trim(rs!conftipo) & " No coincide"
                Else
                    columna_p1(13, 0, 0) = columna_p1(13, 0, 0) & "," & rs!confval
                End If
            Case "35"
                If columna_p1(14, 0, 0) = "" Then
                    columna_p1(14, 0, 0) = Trim(rs!confval)
                    columna_p1(14, 1, 0) = Trim(rs!conftipo)
                    columna_p1(14, 1, 1) = Trim(rs!confetiq)
                    confval2_p1(14) = Trim(rs!confval2)
                ElseIf columna_p1(14, 1, 0) <> Trim(rs!conftipo) Then
                    Flog.writeline "Error de tipo de datos en columna N°: 35 " & Trim(rs!conftipo) & " No coincide"
                Else
                    columna_p1(14, 0, 0) = columna_p1(14, 0, 0) & "," & rs!confval
                End If
            Case "36"
                If columna_p1(15, 0, 0) = "" Then
                    columna_p1(15, 0, 0) = Trim(rs!confval)
                    columna_p1(15, 1, 0) = Trim(rs!conftipo)
                    columna_p1(15, 1, 1) = Trim(rs!confetiq)
                    confval2_p1(15) = Trim(rs!confval2)
                ElseIf columna_p1(15, 1, 0) <> Trim(rs!conftipo) Then
                    Flog.writeline "Error de tipo de datos en columna N°: 36 " & Trim(rs!conftipo) & " No coincide"
                Else
                    columna_p1(15, 0, 0) = columna_p1(15, 0, 0) & "," & rs!confval
                End If
            Case "37"
                If columna_p1(16, 0, 0) = "" Then
                    columna_p1(16, 0, 0) = Trim(rs!confval)
                    columna_p1(16, 1, 0) = Trim(rs!conftipo)
                    columna_p1(16, 1, 1) = Trim(rs!confetiq)
                    confval2_p1(16) = Trim(rs!confval2)
                ElseIf columna_p1(16, 1, 0) <> Trim(rs!conftipo) Then
                    Flog.writeline "Error de tipo de datos en columna N°: 37 " & Trim(rs!conftipo) & " No coincide"
                Else
                    columna_p1(16, 0, 0) = columna_p1(16, 0, 0) & "," & rs!confval
                End If
            Case "38"
                If columna_p1(17, 0, 0) = "" Then
                    columna_p1(17, 0, 0) = Trim(rs!confval)
                    columna_p1(17, 1, 0) = Trim(rs!conftipo)
                    columna_p1(17, 1, 1) = Trim(rs!confetiq)
                    confval2_p1(17) = Trim(rs!confval2)
                ElseIf columna_p1(17, 1, 0) <> Trim(rs!conftipo) Then
                    Flog.writeline "Error de tipo de datos en columna N°: 38 " & Trim(rs!conftipo) & " No coincide"
                Else
                    columna_p1(17, 0, 0) = columna_p1(17, 0, 0) & "," & rs!confval
                End If
            Case "39"
                If columna_p1(18, 0, 0) = "" Then
                    columna_p1(18, 0, 0) = Trim(rs!confval)
                    columna_p1(18, 1, 0) = Trim(rs!conftipo)
                    columna_p1(18, 1, 1) = Trim(rs!confetiq)
                    confval2_p1(18) = Trim(rs!confval2)
                ElseIf columna_p1(18, 1, 0) <> Trim(rs!conftipo) Then
                    Flog.writeline "Error de tipo de datos en columna  N°: 39 " & Trim(rs!conftipo) & " No coincide"
                Else
                    columna_p1(18, 0, 0) = columna_p1(18, 0, 0) & "," & rs!confval
                End If
            Case "40"
                If columna_p1(19, 0, 0) = "" Then
                    columna_p1(19, 0, 0) = Trim(rs!confval)
                    columna_p1(19, 1, 0) = Trim(rs!conftipo)
                    columna_p1(19, 1, 1) = Trim(rs!confetiq)
                    confval2_p1(19) = Trim(rs!confval2)
                ElseIf columna_p1(19, 1, 0) <> Trim(rs!conftipo) Then
                    Flog.writeline "Error de tipo de datos en columna  N°: 40 " & Trim(rs!conftipo) & " No coincide"
                Else
                    columna_p1(19, 0, 0) = columna_p1(19, 0, 0) & "," & rs!confval
                End If


    
         End Select
         
        rs.MoveNext
    Loop
    
End If
rs.Close




'------------------------------------------
'SELECCION DE EMPLEADOS SEGUN FILTRO
'------------------------------------------

StrSql2 = ""
'___________________
'TODOS LOS EMPLEADOS
If Empleados = "-1" Then
    StrSql2 = " AND empleado.ternro not in (SELECT ternro from his_estructura agencia "
    StrSql2 = StrSql2 & " WHERE agencia.tenro=28 AND (agencia.htetdesde<=" & ConvFecha(Fecha)
    StrSql2 = StrSql2 & " AND (agencia.htethasta is null or agencia.htethasta>=" & ConvFecha(Fecha) & ")) )"
Else
    If Empleados = "-2" Then
        StrSql2 = " AND empleado.ternro in (SELECT ternro from his_estructura agencia "
        StrSql2 = StrSql2 & " WHERE agencia.tenro= 28 AND (agencia.htetdesde<=" & ConvFecha(Fecha)
        StrSql2 = StrSql2 & " AND (agencia.htethasta is null or agencia.htethasta>=" & ConvFecha(Fecha) & " )) )"
    Else
        '______________________________
        'CUANDO SE SELECCIONA X AGENCIA
        If Empleados <> "0" Then
            StrSql2 = " AND empleado.ternro in (SELECT ternro from his_estructura agencia "
            StrSql2 = StrSql2 & " WHERE agencia.tenro=28 and agencia.estrnro=" & Empleados
            StrSql2 = StrSql2 & "  AND (agencia.htetdesde<=" & ConvFecha(Fecha)
            StrSql2 = StrSql2 & " AND (agencia.htethasta is null or agencia.htethasta>=" & ConvFecha(Fecha) & ")) )"
        End If
    End If
End If

'____________________________________
'SOLO CUANDO SE SELECCIONAN 3 NIVELES

If t_nivel3 <> "" And t_nivel3 <> "0" Then
    StrSql = "SELECT DISTINCT empleado.ternro,empleg, terape, terape2, ternom, estact1.tenro  tenro1, estact1.estrnro  estrnro1, "
    StrSql = StrSql & " estact2.tenro  tenro2, estact2.estrnro  estrnro2, estact3.tenro  tenro3, estact3.estrnro  estrnro3 "
    StrSql = StrSql & ", estructura1.estrdabr  estrdabr1 , estructura2.estrdabr  estrdabr2, estructura3.estrdabr  estrdabr3 "
    StrSql = StrSql & " FROM empleado "
    StrSql = StrSql & " INNER JOIN his_estructura estact1 ON empleado.ternro = estact1.ternro  AND estact1.tenro  = " & t_nivel1
    StrSql = StrSql & " AND (estact1.htetdesde<=" & ConvFecha(Fecha) & " AND (estact1.htethasta is null or estact1.htethasta>=" & ConvFecha(Fecha) & "))"
    '_______________________________
    'SE LE ASIGNA 1 VALOR AL NIVEL 1
    If t_estruc1 <> "" And t_estruc1 <> "0" Then
        StrSql = StrSql & " AND estact1.estrnro =" & t_estruc1
    End If
    StrSql = StrSql & " INNER JOIN estructura estructura1 ON estructura1.estrnro=estact1.estrnro "
    StrSql = StrSql & " INNER JOIN his_estructura estact2 ON empleado.ternro = estact2.ternro  AND estact2.tenro  = " & t_nivel2
    StrSql = StrSql & " AND (estact2.htetdesde<=" & ConvFecha(Fecha) & " AND (estact2.htethasta is null or estact2.htethasta>=" & ConvFecha(Fecha) & "))"
    
    '_______________________________
    'SE LE ASIGNA 1 VALOR AL NIVEL 2
    If t_estruc2 <> "" And t_estruc2 <> "0" Then
        StrSql = StrSql & " AND estact2.estrnro =" & t_estruc2
    End If
    
    StrSql = StrSql & " INNER JOIN estructura estructura2 ON estructura2.estrnro=estact2.estrnro "
    StrSql = StrSql & " INNER JOIN his_estructura estact3 ON empleado.ternro = estact3.ternro  AND estact3.tenro  = " & t_nivel3
    StrSql = StrSql & " AND (estact3.htetdesde<=" & ConvFecha(Fecha) & " AND (estact3.htethasta is null or estact3.htethasta>=" & ConvFecha(Fecha) & "))"
    
    '_______________________________
    'SE LE ASIGNA 1 VALOR AL NIVEL 3
    If t_estruc3 <> "" And t_estruc3 <> "0" Then
        StrSql = StrSql & " AND estact3.estrnro =" & t_estruc3
    End If
    
    StrSql = StrSql & " INNER JOIN estructura estructura3 ON estructura3.estrnro=estact3.estrnro "
    StrSql = StrSql & " WHERE " & Estado & StrSql2
    StrSql = StrSql & " ORDER BY tenro1,estrdabr1,tenro2,estrdabr2,tenro3,estrdabr3," & Orden & " " & Ordenado
'_______________________________________
'CUANDO SE SELECCIONA HASTA EL 2DO NIVEL
ElseIf t_nivel2 <> "" And t_nivel2 <> "0" Then
        StrSql = "SELECT DISTINCT empleado.ternro,empleg, terape, terape2, ternom, estact1.tenro  tenro1, estact1.estrnro  estrnro1, "
        StrSql = StrSql & " estact2.tenro  tenro2, estact2.estrnro  estrnro2 "
        StrSql = StrSql & ", estructura1.estrdabr  estrdabr1 , estructura2.estrdabr  estrdabr2 "
        StrSql = StrSql & " FROM empleado "
        StrSql = StrSql & " INNER JOIN his_estructura estact1 ON empleado.ternro = estact1.ternro  AND estact1.tenro  = " & t_nivel1
        StrSql = StrSql & " AND (estact1.htetdesde<=" & ConvFecha(Fecha) & " AND (estact1.htethasta is null or estact1.htethasta>=" & ConvFecha(Fecha) & "))"
        
        If t_estruc1 <> "" And t_estruc1 <> "0" Then
            StrSql = StrSql & " AND estact1.estrnro =" & t_estruc1
        End If
        
        StrSql = StrSql & " INNER JOIN estructura estructura1 ON estructura1.estrnro=estact1.estrnro "
        StrSql = StrSql & " INNER JOIN his_estructura estact2 ON empleado.ternro = estact2.ternro  AND estact2.tenro  = " & t_nivel2
        StrSql = StrSql & " AND (estact2.htetdesde<=" & ConvFecha(Fecha) & " AND (estact2.htethasta is null or estact2.htethasta>=" & ConvFecha(Fecha) & "))"
        
        If t_estruc2 <> "" And t_estruc2 <> "0" Then
            StrSql = StrSql & " AND estact2.estrnro =" & t_estruc2
        End If
        
        StrSql = StrSql & " INNER JOIN estructura estructura2 ON estructura2.estrnro=estact2.estrnro "
        StrSql = StrSql & " WHERE " & Estado & StrSql2
        StrSql = StrSql & " ORDER BY tenro1,estrdabr1,tenro2,estrdabr2," & Orden & " " & Ordenado
'______________________________________
'CUANDO SOLO SE SELECCIONA EL 1ER NIVEL
ElseIf t_nivel1 <> "" And t_nivel1 <> "0" Then
        StrSql = " SELECT DISTINCT empleado.ternro,empleg, terape, terape2, ternom, estact1.tenro  tenro1, estact1.estrnro  estrnro1 "
        StrSql = StrSql & ", estructura1.estrdabr  estrdabr1  "
        StrSql = StrSql & " FROM empleado "
        StrSql = StrSql & " INNER JOIN his_estructura estact1 ON empleado.ternro = estact1.ternro  AND estact1.tenro  = " & t_nivel1
        StrSql = StrSql & " AND (estact1.htetdesde<=" & ConvFecha(Fecha) & " AND (estact1.htethasta is null or estact1.htethasta>=" & ConvFecha(Fecha) & "))"
        If t_estruc1 <> "" And t_estruc1 <> "0" Then
            StrSql = StrSql & " AND estact1.estrnro =" & t_estruc1
        End If
        
        StrSql = StrSql & " INNER JOIN estructura estructura1 ON estructura1.estrnro=estact1.estrnro "
        StrSql = StrSql & " WHERE " & Estado & StrSql2
        StrSql = StrSql & " ORDER BY tenro1,estrdabr1," & Orden & " " & Ordenado
'______________________________________________
'CUANDO NO HAY NIVEL DE ESTRUCTURA SELECCIONADO
Else
    StrSql = "SELECT DISTINCT empleado.ternro,empleg, terape, terape2, ternom "
    StrSql = StrSql & " FROM empleado "
    StrSql = StrSql & " WHERE " & Estado & StrSql2
    StrSql = StrSql & " ORDER BY " & Orden & " " & Ordenado
     '     End If
    'End If
End If

OpenRecordset StrSql, rs

Select_empleados = StrSql
If rs.EOF Then
    Flog.writeline "No se encontraron empleados"
    Exit Sub
Else

    'Determino la proporcion de progreso
    Progreso = 0
    CEmpleadosAProc = rs.RecordCount
    If CEmpleadosAProc = 0 Then
        CEmpleadosAProc = 1
    End If
    IncPorc = (50 / CEmpleadosAProc)
    '¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
    'INSERTO CABECERA
    '¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
    Flog.writeline "Preparo insertar la cabecera"
    StrSql = "INSERT into rep_list_asist_dia_cab "
    StrSql = StrSql & "(bpronro,fecha,hora,iduser,empleados) "
    StrSql = StrSql & " VALUES (" & bpronro & "," & ConvFecha(Date) & ",'" & Format_StrNro(Hour(Now), 2, True, 0) & Format_StrNro(Minute(Now), 2, True, 0) & "','" & usuario & "',0)"
    objConn.Execute StrSql, , adExecuteNoRecords
    
    'BUSCO ULT. REGISTRO INSERTADO
    repnro = getLastIdentity(objConn, "rep_list_asist_dia_cab")
    '¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
    Flog.writeline "Ult. registro insertado en la tabla de cabecera: " & repnro
    
    '______________________________________________
    'RECORRO LOS EMPLEADOS SELECTADOS POR EL FILTRO
    Do While Not rs.EOF
    
    Progreso = Progreso + IncPorc
    'Guardo cantidad de empleados selectados según filtro
    cant_empleados = cant_empleados + 1
    Ternro_empleados = Ternro_empleados & "," & rs!Ternro
       
       'nombre = rs!ternom & " " & rs!ternom2
       'Apellido = rs!terape & " " & rs!terape2
       
       'Guardo nombre y apellido
       nombre = rs!ternom
       Apellido = rs!terape
        
        '_____________________________________________________________________________
        'DEVUELVE N° DE ESTRUCTURA Y DESCRIPCION DEL EMPLEADO SEGUN TIPO E. | PARTE 2
        If testrnro1 <> 0 Then
            estrnro1 = estructura_actual(testrnro1, rs!empleg, Fecha)
            estrnro1_desc = nombre_estructura(estrnro1)
        End If
        If testrnro2 <> 0 Then
            estrnro2 = estructura_actual(testrnro2, rs!empleg, Fecha)
            estrnro2_desc = nombre_estructura(estrnro2)
        End If
        If testrnro3 <> 0 Then
            estrnro3 = estructura_actual(testrnro3, rs!empleg, Fecha)
            estrnro3_desc = nombre_estructura(estrnro3)
        End If
        If testrnro4 <> 0 Then
           estrnro4 = estructura_actual(testrnro4, rs!empleg, Fecha)
            estrnro4_desc = nombre_estructura(estrnro4)
        End If
        If testrnro5 <> 0 Then
            estrnro5 = estructura_actual(testrnro5, rs!empleg, Fecha)
            estrnro5_desc = nombre_estructura(estrnro5)
        End If
        '----------------------------------------------------------------
        
        '¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
        'PREPARO INSERT
        '¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
        StrSql_ins = "INSERT INTO rep_list_asist_dia_p2 "
        StrSql_ins = StrSql_ins & "("
        StrSql_ins = StrSql_ins & "repnro,ternro"
        StrSql_ins = StrSql_ins & ",nombre,apellido"
        StrSql_ins = StrSql_ins & ",te1,tedesc1,estrnro1,estrdesc1"
        StrSql_ins = StrSql_ins & ",te2,tedesc2,estrnro2,estrdesc2"
        StrSql_ins = StrSql_ins & ",te3,tedesc3,estrnro3,estrdesc3"
        StrSql_ins = StrSql_ins & ",te4,tedesc4,estrnro4,estrdesc4"
        StrSql_ins = StrSql_ins & ",te5,tedesc5,estrnro5,estrdesc5 "
        StrSql_ins = StrSql_ins & ",col1,col1_desc"
        StrSql_ins = StrSql_ins & ",col2,col2_desc"
        StrSql_ins = StrSql_ins & ",col3,col3_desc"
        StrSql_ins = StrSql_ins & ",col4,col4_desc"
        StrSql_ins = StrSql_ins & ",col5,col5_desc"
        StrSql_ins = StrSql_ins & ",col6,col6_desc"
        StrSql_ins = StrSql_ins & ",col7,col7_desc"
        StrSql_ins = StrSql_ins & ",col8,col8_desc"
        StrSql_ins = StrSql_ins & ",col9,col9_desc"
        StrSql_ins = StrSql_ins & ",col10,col10_desc"
        StrSql_ins = StrSql_ins & ",col11,col11_desc"
        StrSql_ins = StrSql_ins & ",col12,col12_desc"
        StrSql_ins = StrSql_ins & ",col13,col13_desc"
        StrSql_ins = StrSql_ins & ",col14,col14_desc"
        StrSql_ins = StrSql_ins & ",col15,col15_desc"
        StrSql_ins = StrSql_ins & ",col16,col16_desc"
        StrSql_ins = StrSql_ins & ",col17,col17_desc"
        StrSql_ins = StrSql_ins & ",col18,col18_desc"
        StrSql_ins = StrSql_ins & ",col19,col19_desc"
        StrSql_ins = StrSql_ins & ",col20,col20_desc"
        StrSql_ins = StrSql_ins & ",obs,acum_anual,acum_total"
        StrSql_ins = StrSql_ins & ")"
        StrSql_ins = StrSql_ins & " VALUES "
  
        '_________________________________________________________________
        'RECORRE EL ARRAY columna() Y EJECUTA FUNCION CORRESPONDIENTE
        For nro_col = 0 To UBound(columna)
            
            'Formateo en NULL
            If col(nro_col) = "" Then
                col(nro_col) = "Null"
            End If
            
            If columna(nro_col, 0, 0) <> "" Then
            '¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
                'CALCULA TIPO DE HORAS
            '¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
                If columna(nro_col, 1, 0) <> "" And columna(nro_col, 1, 0) = "TH" Then
                   'SI CALCULA TIPO DE HORAS
                    col(nro_col) = tipodehora(rs!Ternro, columna(nro_col, 0, 0), Fecha)
                                
                    'If col(nro_col) <> 0 Then
                    'ANTES DEL INSERT
                    
                    'If col(nro_col) <> 0 Or confval2(nro_col) = "REG" Then
                    If col(nro_col) <> 0 Then
                        hayreg = "-1"
                        'If (confval2(nro_col)) <> ""  Then
                        'If Not IsNumeric(col(nro_col)) Or col(nro_col) = 0 Then
                        If Not IsNumeric(col(nro_col)) Then
                            If confval2(nro_col) = "REG" Then
                                'Devuelve X si no hay registracion
                                hayreg = hayregistracion(rs!Ternro, Fecha)
                                
                            End If
                            
                            
                            
                            'If hayreg = "-1" Then
                                'Si hay registracion
                            '    columna(nro_col, 1, 1) = columna(nro_col, 1, 1) & "@" & Replace(col(nro_col), "'", "")
                            'Else
                                'Si no hay registracion
                            If hayreg = "X" Then
                                columna(nro_col, 1, 1) = columna(nro_col, 1, 1) & "@" & hayreg
                                'col(nro_col) = "Null"
                            ElseIf hayreg = "" Then
                                columna(nro_col, 1, 1) = columna(nro_col, 1, 1)
                                'col(nro_col) = "Null"
                            Else
                                columna(nro_col, 1, 1) = columna(nro_col, 1, 1) & "@" & Replace(col(nro_col), "'", "")
                            End If
                            'ElseIf col(nro_col) <> 0 Then
                            '    columna(nro_col, 1, 1) = columna(nro_col, 1, 1) & "@" & Replace(col(nro_col), "'", "")
                                
                            'End If
                            
                            col(nro_col) = "Null"
                            'Si tiene registracion muestra horas
                                'columna(nro_col, 1, 1) = columna(nro_col, 1, 1) & "@" & Replace(col(nro_col), "'", "")
                                'col(nro_col) = "NULL"
                           
                            

                            tipo = "TH"
                            acumulado = tipodehora_acum(rs!Ternro, columna(nro_col, 0, 0), Fecha, tipo)
                        End If
                        
                       ' If Not IsNumeric(col(nro_col)) Then
                            'Texto_aux = Replace(col(nro_col), "'", "")
                            'col(nro_col) = "NULL"
                            
                        'End If
                        
                        'FORMATEO PARA INSERTARLO COMO NULL
                        'Flog.writeline "col: " & col(nro_col)
                        'If col(nro_col) = 0 Then
                        '  col(nro_col) = "NULL"
                        'End If
                        
                        'If (hayreg = "X") Or (hayreg = "" And col(nro_col) <> 0) Then
                        If columna(nro_col, 1, 1) = "Falta" And hayreg = "" Then
                            'NO INSERTA REGISTRO SOLO CUANDO ES DE TIPO FALTA Y TIEN REGISTRACIONES.
                            Flog.writeline
                            Flog.writeline "NO SE INSERTA REGISTRO CUANDO ES DE TIPO FALTA Y TIENE REGISTRACIONES, ES DECIR NO TIENE ANORMALIDADES!"
                            Flog.writeline
                        Else
                            'INSERTA SOLO SI NO HUBO REGISTRACION...ES DECIR SI EN hayreg SE GUARDA UNA X
                            'INSERTO____________________________________________
                            StrSql = StrSql_ins
                            StrSql = StrSql & "("
                            StrSql = StrSql & repnro & "," & rs!Ternro & ",'" & nombre & "','" & Apellido & "'"
                            StrSql = StrSql & "," & testrnro1 & ",'" & col_nombre1 & "'," & estrnro1 & ",'" & estrnro1_desc & "'"
                            StrSql = StrSql & "," & testrnro2 & ",'" & col_nombre2 & "'," & estrnro2 & ",'" & estrnro2_desc & "'"
                            StrSql = StrSql & "," & testrnro3 & ",'" & col_nombre3 & "'," & estrnro3 & ",'" & estrnro3_desc & "'"
                            StrSql = StrSql & "," & testrnro4 & ",'" & col_nombre4 & "'," & estrnro4 & ",'" & estrnro4_desc & "'"
                            StrSql = StrSql & "," & testrnro5 & ",'" & col_nombre5 & "'," & estrnro5 & ",'" & estrnro5_desc & "'"
                            StrSql = StrSql & "," & col(0) & ",'" & columna(0, 1, 1) & "'"
                            StrSql = StrSql & "," & col(1) & ",'" & columna(1, 1, 1) & "'"
                            StrSql = StrSql & "," & col(2) & ",'" & columna(2, 1, 1) & "'"
                            StrSql = StrSql & "," & col(3) & ",'" & columna(3, 1, 1) & "'"
                            StrSql = StrSql & "," & col(4) & ",'" & columna(4, 1, 1) & "'"
                            StrSql = StrSql & "," & col(5) & ",'" & columna(5, 1, 1) & "'"
                            StrSql = StrSql & "," & col(6) & ",'" & columna(6, 1, 1) & "'"
                            StrSql = StrSql & "," & col(7) & ",'" & columna(7, 1, 1) & "'"
                            StrSql = StrSql & "," & col(8) & ",'" & columna(8, 1, 1) & "'"
                            StrSql = StrSql & "," & col(9) & ",'" & columna(9, 1, 1) & "'"
                            StrSql = StrSql & "," & col(10) & ",'" & columna(10, 1, 1) & "'"
                            StrSql = StrSql & "," & col(11) & ",'" & columna(11, 1, 1) & "'"
                            StrSql = StrSql & "," & col(12) & ",'" & columna(12, 1, 1) & "'"
                            StrSql = StrSql & "," & col(13) & ",'" & columna(13, 1, 1) & "'"
                            StrSql = StrSql & "," & col(14) & ",'" & columna(14, 1, 1) & "'"
                            StrSql = StrSql & "," & col(15) & ",'" & columna(15, 1, 1) & "'"
                            StrSql = StrSql & "," & col(16) & ",'" & columna(16, 1, 1) & "'"
                            StrSql = StrSql & "," & col(17) & ",'" & columna(17, 1, 1) & "'"
                            StrSql = StrSql & "," & col(18) & ",'" & columna(18, 1, 1) & "'"
                            StrSql = StrSql & "," & col(19) & ",'" & columna(19, 1, 1) & "'"
                            StrSql = StrSql & ",'" & observaciones & "' , " & acumulado & ", NULL"
                            
                            StrSql = StrSql & ")"
                            
                            'EJECUTO
                            'StrSql = StrSql
                            objConn.Execute StrSql, , adExecuteNoRecords
                            Flog.writeline "1er insert en PARTE 2"
                        End If
                            'If (confval2(nro_col)) <> "" Then
                            If col(nro_col) <> "" Then
                                col(nro_col) = "Null"
                                'Restauro el valor original de la columna
                                'columna(nro_col, 1, 1) = Trim(Replace(columna(nro_col, 1, 1), "@" & confval2(nro_col), ""))
                                
                            End If
                        
              
                        'End If
                        If hayreg <> "" Then
                                columna(nro_col, 1, 1) = Left(columna(nro_col, 1, 1), InStr(1, columna(nro_col, 1, 1), "@") - 1)
                        End If
                            
                        'columna(nro_col, 1, 1) = columna(nro_col, 1, 1) & "@" & Replace(col(nro_col), "'", "")
                       'columna(nro_col, 1, 1) = "Null"
                        'columna(nro_col, 1, 1) = Trim(Replace(columna(nro_col, 1, 1), "@" & Trim(Mid(licencia_anual(0), 1, pos1 - 1)), ""))
                        Texto_aux = ""
                        'Reseteo en NULL
                        'col(nro_col) = "Null"
                        
                    End If
                    
                    col(nro_col) = "Null"
                    
                ElseIf columna(nro_col, 1, 0) <> "" And columna(nro_col, 1, 0) = "LIC" Then
                'SI CALCULA LICENCIAS
                    If IsNull(confval2(nro_col)) Then
                        confval2(nro_col) = 0
                    End If
                    
                    licencia_anual = Split(lic_anual(rs!Ternro, columna(nro_col, 0, 0), confval2(nro_col), Fecha), "@")
                    
                    If UBound(licencia_anual) > -1 Then
                                        
                        pos1 = InStr(1, licencia_anual(0), "-", 1)
                        
                        If confval2(nro_col) <> "" And confval2(nro_col) <> 0 Then
                            columna(nro_col, 1, 1) = columna(nro_col, 1, 1) & "@" & confval2(nro_col)
                        Else
                            columna(nro_col, 1, 1) = columna(nro_col, 1, 1) & "@" & Trim(Mid(licencia_anual(0), 1, pos1 - 1))
                        End If
                        
                        observaciones = licencia_anual(0)
                        acumulado = licencia_anual(1)
                        acumulado_total = licencia_anual(2)
                                                                        
                        'INSERTO____________________________________________
                        StrSql = StrSql_ins
                        StrSql = StrSql & "("
                        StrSql = StrSql & repnro & "," & rs!Ternro & ",'" & nombre & "','" & Apellido & "'"
                        StrSql = StrSql & "," & testrnro1 & ",'" & col_nombre1 & "'," & estrnro1 & ",'" & estrnro1_desc & "'"
                        StrSql = StrSql & "," & testrnro2 & ",'" & col_nombre2 & "'," & estrnro2 & ",'" & estrnro2_desc & "'"
                        StrSql = StrSql & "," & testrnro3 & ",'" & col_nombre3 & "'," & estrnro3 & ",'" & estrnro3_desc & "'"
                        StrSql = StrSql & "," & testrnro4 & ",'" & col_nombre4 & "'," & estrnro4 & ",'" & estrnro4_desc & "'"
                        StrSql = StrSql & "," & testrnro5 & ",'" & col_nombre5 & "'," & estrnro5 & ",'" & estrnro5_desc & "'"
                        StrSql = StrSql & "," & col(0) & ",'" & columna(0, 1, 1) & "'"
                        StrSql = StrSql & "," & col(1) & ",'" & columna(1, 1, 1) & "'"
                        StrSql = StrSql & "," & col(2) & ",'" & columna(2, 1, 1) & "'"
                        StrSql = StrSql & "," & col(3) & ",'" & columna(3, 1, 1) & "'"
                        StrSql = StrSql & "," & col(4) & ",'" & columna(4, 1, 1) & "'"
                        StrSql = StrSql & "," & col(5) & ",'" & columna(5, 1, 1) & "'"
                        StrSql = StrSql & "," & col(6) & ",'" & columna(6, 1, 1) & "'"
                        StrSql = StrSql & "," & col(7) & ",'" & columna(7, 1, 1) & "'"
                        StrSql = StrSql & "," & col(8) & ",'" & columna(8, 1, 1) & "'"
                        StrSql = StrSql & "," & col(9) & ",'" & columna(9, 1, 1) & "'"
                        StrSql = StrSql & "," & col(10) & ",'" & columna(10, 1, 1) & "'"
                        StrSql = StrSql & "," & col(11) & ",'" & columna(11, 1, 1) & "'"
                        StrSql = StrSql & "," & col(12) & ",'" & columna(12, 1, 1) & "'"
                        StrSql = StrSql & "," & col(13) & ",'" & columna(13, 1, 1) & "'"
                        StrSql = StrSql & "," & col(14) & ",'" & columna(14, 1, 1) & "'"
                        StrSql = StrSql & "," & col(15) & ",'" & columna(15, 1, 1) & "'"
                        StrSql = StrSql & "," & col(16) & ",'" & columna(16, 1, 1) & "'"
                        StrSql = StrSql & "," & col(17) & ",'" & columna(17, 1, 1) & "'"
                        StrSql = StrSql & "," & col(18) & ",'" & columna(18, 1, 1) & "'"
                        StrSql = StrSql & "," & col(19) & ",'" & columna(19, 1, 1) & "'"
                        StrSql = StrSql & ",'" & observaciones & "' , " & acumulado & "," & acumulado_total & ""
                        StrSql = StrSql & ")"
                        objConn.Execute StrSql, , adExecuteNoRecords
                        Flog.writeline "2do insert en PARTE 2"
                    
                        'Restauro columna
                        If confval2(nro_col) <> "VAC" Then
                            columna(nro_col, 1, 1) = Trim(Replace(columna(nro_col, 1, 1), "@" & Trim(Mid(licencia_anual(0), 1, pos1 - 1)), ""))
                        Else
                            columna(nro_col, 1, 1) = Trim(Replace(columna(nro_col, 1, 1), "@" & confval2(nro_col), ""))
                        End If
                        pos1 = 0
                        observaciones = ""
                        acumulado = "NULL"
                        acumulado_total = "NULL"
                    End If
                
                
                ElseIf columna(nro_col, 1, 0) <> "" And columna(nro_col, 1, 0) = "NOV" Then
                'SI CALCULA NOVEDADES
                    
                    novedad_aux = novedad(rs!Ternro, columna(nro_col, 0, 0), Fecha)
                    
                    If novedad_aux <> "" Then
                        novedad_aux = Split(novedad(rs!Ternro, columna(nro_col, 0, 0), Fecha), "@")
                        observaciones = novedad_aux(0)
                    End If
                    
                    If observaciones <> "" Then
                        
                        'Si no tiene configurado confval2
                        
                        If confval2(nro_col) = "" Then
                            acumulado = tipodehora(rs!Ternro, novedad_aux(1), Fecha)
                            If acumulado_total = "" Then
                                acumulado_total = "NULL"
                            End If
                        Else
                            'acumulado = "NULL"
                            acumulado_total = "NULL"
                            columna(nro_col, 1, 1) = columna(nro_col, 1, 1) & "@" & confval2(nro_col)
                        End If
                        'acumulado_total = tipodehora_acum(rs!Ternro, novedad_aux(1), Fecha)
                        acumulado = tipodehora_acum(rs!Ternro, novedad_aux(1), Fecha, tipo)
                'INSERTO____________________________________________
                        StrSql = StrSql_ins
                        StrSql = StrSql & "("
                        StrSql = StrSql & repnro & "," & rs!Ternro & ",'" & nombre & "','" & Apellido & "'"
                        StrSql = StrSql & "," & testrnro1 & ",'" & col_nombre1 & "'," & estrnro1 & ",'" & estrnro1_desc & "'"
                        StrSql = StrSql & "," & testrnro2 & ",'" & col_nombre2 & "'," & estrnro2 & ",'" & estrnro2_desc & "'"
                        StrSql = StrSql & "," & testrnro3 & ",'" & col_nombre3 & "'," & estrnro3 & ",'" & estrnro3_desc & "'"
                        StrSql = StrSql & "," & testrnro4 & ",'" & col_nombre4 & "'," & estrnro4 & ",'" & estrnro4_desc & "'"
                        StrSql = StrSql & "," & testrnro5 & ",'" & col_nombre5 & "'," & estrnro5 & ",'" & estrnro5_desc & "'"
                        StrSql = StrSql & "," & col(0) & ",'" & columna(0, 1, 1) & "'"
                        StrSql = StrSql & "," & col(1) & ",'" & columna(1, 1, 1) & "'"
                        StrSql = StrSql & "," & col(2) & ",'" & columna(2, 1, 1) & "'"
                        StrSql = StrSql & "," & col(3) & ",'" & columna(3, 1, 1) & "'"
                        StrSql = StrSql & "," & col(4) & ",'" & columna(4, 1, 1) & "'"
                        StrSql = StrSql & "," & col(5) & ",'" & columna(5, 1, 1) & "'"
                        StrSql = StrSql & "," & col(6) & ",'" & columna(6, 1, 1) & "'"
                        StrSql = StrSql & "," & col(7) & ",'" & columna(7, 1, 1) & "'"
                        StrSql = StrSql & "," & col(8) & ",'" & columna(8, 1, 1) & "'"
                        StrSql = StrSql & "," & col(9) & ",'" & columna(9, 1, 1) & "'"
                        StrSql = StrSql & "," & col(10) & ",'" & columna(10, 1, 1) & "'"
                        StrSql = StrSql & "," & col(11) & ",'" & columna(11, 1, 1) & "'"
                        StrSql = StrSql & "," & col(12) & ",'" & columna(12, 1, 1) & "'"
                        StrSql = StrSql & "," & col(13) & ",'" & columna(13, 1, 1) & "'"
                        StrSql = StrSql & "," & col(14) & ",'" & columna(14, 1, 1) & "'"
                        StrSql = StrSql & "," & col(15) & ",'" & columna(15, 1, 1) & "'"
                        StrSql = StrSql & "," & col(16) & ",'" & columna(16, 1, 1) & "'"
                        StrSql = StrSql & "," & col(17) & ",'" & columna(17, 1, 1) & "'"
                        StrSql = StrSql & "," & col(18) & ",'" & columna(18, 1, 1) & "'"
                        StrSql = StrSql & "," & col(19) & ",'" & columna(19, 1, 1) & "'"
                        StrSql = StrSql & ",'" & observaciones & "', " & acumulado & "," & acumulado_total & ""
                        StrSql = StrSql & ")"
                        objConn.Execute StrSql, , adExecuteNoRecords
                        Flog.writeline "3er insert en PARTE 2"
                        
                        'Restauro columna
                        columna(nro_col, 1, 1) = Trim(Replace(columna(nro_col, 1, 1), "@" & confval2(nro_col), ""))
                        observaciones = ""
                        acumulado = 0
                        acumulado_total = 0
                   End If
                End If
            End If
                     
        Next
        '-------------------------------------------------------------------
              
        rs.MoveNext
        'Guardo en batch_proceso
         MyBeginTrans
         StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso
         'StrSql = StrSql & " , bprctiempo = " & TiempoInicialProceso
         StrSql = StrSql & " WHERE bpronro = " & NroProcesoBatch
         objconnProgreso.Execute StrSql, , adExecuteNoRecords
         MyCommitTrans
    Loop
    
   '¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
   'REPORTE PARTE 1
   '¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
   'Recorro columnas configuradas -- PARTE 1
   IncPorc = (25 / 20)
    For c = 0 To 19
        Progreso = Progreso + IncPorc
        
        Flog.writeline "Busco NOV para P1"
        If Trim(confval2_p1(c)) = "P1" And Trim(columna_p1(c, 1, 0)) = "NOV" Then
            total_x_col = novedad_p1(Ternro_empleados, columna_p1(c, 0, 0), Fecha)
        End If
      
        Flog.writeline "Busco LIC para P1"
        If Trim(confval2_p1(c)) = "P1" And Trim(columna_p1(c, 1, 0)) = "LIC" Then
            total_x_col = lic_anual_p1(Ternro_empleados, columna_p1(c, 0, 0), Fecha)
        End If
        
        Flog.writeline "Busco TH para P1"
        If Trim(confval2_p1(c)) = "P1" And Trim(columna_p1(c, 1, 0)) = "TH" Then
            total_x_col = tipodehora_p1(Ternro_empleados, columna_p1(c, 0, 0), Fecha)
        End If
        'Flog.Writeline StrSql
        
        'Inserto
        If Trim(total_x_col) <> "" Then
            If Trim(total_x_col) <> "0@" Then
                total_x_col2 = Split(total_x_col, "@", -1, 1)
                
                descripcion_p1 = Split((columna_p1(c, 1, 1)), "@")
                
                StrSql = "INSERT INTO rep_list_asist_dia_p1 "
                StrSql = StrSql & "(repnro,cod,descripcion,cant,porc)"
                StrSql = StrSql & " VALUES "
                StrSql = StrSql & " (" & repnro & ",'" & descripcion_p1(0) & "','" & descripcion_p1(1) & "'," & CInt(total_x_col2(0)) & "," & Replace(Round((CInt(total_x_col2(0)) / cant_empleados) * 100, 4), ",", ".") & ")"
                objConn.Execute StrSql, , adExecuteNoRecords
                Flog.writeline "Inserto en PARTE 1"
            End If
            
            total_x_col = ""
            
        End If
        
        
        'Guardo en batch_proceso
         MyBeginTrans
         StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso
         'StrSql = StrSql & " , bprctiempo = " & TiempoInicialProceso
         StrSql = StrSql & " WHERE bpronro = " & NroProcesoBatch
         objconnProgreso.Execute StrSql, , adExecuteNoRecords
         MyCommitTrans
    Next
    
    '¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
    'REPORTE PARTE 3
    '¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
    If Select_empleados <> "" And Trim(hs_ausencia) <> "" Then
        StrSql = Select_empleados
        OpenRecordset StrSql, rs
        
        IncPorc = (25 / CEmpleadosAProc)
        Do While Not rs.EOF
            
            Progreso = Progreso + IncPorc
            datos_supervisor = horcumplido(rs!Ternro, hs_ausencia, Fecha)
            If datos_supervisor <> "" Then
               datos_supervisor = Split(datos_supervisor, "@")
               '_____________________________________________________________________________
               'DEVUELVE N° DE ESTRUCTURA Y DESCRIPCION DEL EMPLEADO SEGUN TIPO E. | PARTE 3
                If testrnro6 <> 0 Then
                    estrnro6 = estructura_actual(testrnro6, datos_supervisor(0), Fecha)
                    estrnro6_desc = nombre_estructura(estrnro6)
                End If
                If testrnro7 <> 0 Then
                    estrnro7 = estructura_actual(testrnro7, datos_supervisor(0), Fecha)
                    estrnro7_desc = nombre_estructura(estrnro7)
                End If
                If testrnro8 <> 0 Then
                    estrnro8 = estructura_actual(testrnro8, datos_supervisor(0), Fecha)
                    estrnro8_desc = nombre_estructura(estrnro8)
                End If
                If testrnro9 <> 0 Then
                   estrnro9 = estructura_actual(testrnro9, datos_supervisor(0), Fecha)
                    estrnro9_desc = nombre_estructura(estrnro9)
                End If
                If testrnro10 <> 0 Then
                    estrnro10 = estructura_actual(testrnro10, datos_supervisor(0), Fecha)
                    estrnro10_desc = nombre_estructura(estrnro10)
                End If
                
                
                 StrSql = "SELECT * FROM rep_list_asist_dia_p3 "
                 StrSql = StrSql & " WHERE repnro = " & repnro
                 StrSql = StrSql & " AND ternro = " & datos_supervisor(0)
                 OpenRecordset StrSql, rs_supervisor
                 If rs_supervisor.EOF Then
                  
                    StrSql = "INSERT INTO rep_list_asist_dia_p3 (repnro,ternro,te1,tedesc1,estrnro1,estrdesc1,te2,tedesc2,estrnro2,estrdesc2"
                    StrSql = StrSql & ",te3,tedesc3,estrnro3,estrdesc3,te4,tedesc4,estrnro4,estrdesc4"
                    StrSql = StrSql & ",te5,tedesc5,estrnro5,estrdesc5,nombre,apellido) "
                    StrSql = StrSql & " VALUES (" & repnro & "," & datos_supervisor(0)
                    StrSql = StrSql & "," & testrnro6 & ",'" & col_nombre6 & "'"
                    StrSql = StrSql & "," & estrnro6 & ",'" & estrnro6_desc & "'"
                    StrSql = StrSql & "," & testrnro7 & ",'" & col_nombre7 & "'"
                    StrSql = StrSql & "," & estrnro7 & ",'" & estrnro7_desc & "'"
                    StrSql = StrSql & "," & testrnro8 & ",'" & col_nombre8 & "'"
                    StrSql = StrSql & "," & estrnro8 & ",'" & estrnro8_desc & "'"
                    StrSql = StrSql & "," & testrnro9 & ",'" & col_nombre9 & "'"
                    StrSql = StrSql & "," & estrnro9 & ",'" & estrnro9_desc & "'"
                    StrSql = StrSql & "," & testrnro10 & ",'" & col_nombre10 & "'"
                    StrSql = StrSql & "," & estrnro10 & ",'" & estrnro10_desc & "'"
                    StrSql = StrSql & ",'" & datos_supervisor(1) & "','" & datos_supervisor(2) & "'"
                    StrSql = StrSql & ")"
                    objConn.Execute StrSql, , adExecuteNoRecords
                    Flog.writeline "Inserto en PARTE 3"
                 End If
                 rs_supervisor.Close
            End If
            rs.MoveNext
        
         MyBeginTrans
         StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso
         'StrSql = StrSql & " , bprctiempo = " & TiempoInicialProceso
         StrSql = StrSql & " WHERE bpronro = " & NroProcesoBatch
         objconnProgreso.Execute StrSql, , adExecuteNoRecords
         MyCommitTrans
         
        Loop
        rs.Close
               
    Else
        Flog.writeline "Es posible que no se encuentre configurado el TH para Ausencias"
    
    End If
        

End If

'¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
'UPDATEO CABECERA CON TOTAL DE EMPLEADOS PROCESADOS
'¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
StrSql = "UPDATE rep_list_asist_dia_cab  set empleados = " & cant_empleados
StrSql = StrSql & " WHERE repnro = " & repnro
objConn.Execute StrSql, , adExecuteNoRecords

'¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
'ELIMINO REGISTROS SEGÚN HISTORICO CONFIGURADO (10 es el minimo permitido)
'¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
If historico < 10 Then
    historico = 10
Else
    StrSql = "SELECT repnro FROM rep_list_asist_dia_cab"
    StrSql = StrSql & " ORDER BY repnro DESC"
    rs.MaxRecords = historico
    OpenRecordset StrSql, rs
    If Not rs.EOF Then
        Do While Not rs.EOF
            'GUARDO SIEMPRE EL ULTIMO REGISTRO
            repnro = rs!repnro
            rs.MoveNext
        Loop
 
    End If
rs.Close

    'COMIENZA LA ELIMINACION
    StrSql = "DELETE FROM  rep_list_asist_dia_cab "
    StrSql = StrSql & " WHERE repnro < " & repnro
    objConn.Execute StrSql, , adExecuteNoRecords
    '---------------------------------------------
    StrSql = "DELETE FROM  rep_list_asist_dia_p1 "
    StrSql = StrSql & " WHERE repnro < " & repnro
    objConn.Execute StrSql, , adExecuteNoRecords
    '---------------------------------------------
    StrSql = "DELETE FROM rep_list_asist_dia_p2 "
    StrSql = StrSql & " WHERE repnro < " & repnro
    objConn.Execute StrSql, , adExecuteNoRecords
    '---------------------------------------------
    StrSql = "DELETE FROM rep_list_asist_dia_p3"
    StrSql = StrSql & " WHERE repnro < " & repnro
    objConn.Execute StrSql, , adExecuteNoRecords
    '---------------------------------------------
End If

Flog.writeline "Total Empleados procesados:" & cant_empleados

End Sub













