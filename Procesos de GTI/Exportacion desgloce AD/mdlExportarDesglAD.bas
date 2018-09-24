Attribute VB_Name = "mdlExportarDesglAD"
Option Explicit

Global Const Version = "1.01"
Global Const FechaModificacion = "02/12/2015" ' Miriam Ruiz   -CAS-34342 - PIRAMIDE - Reporte desglose de Parte de Movilidad
Global Const UltimaModificacion = "Se agregó encriptación - Se cambia el directyorio de exportación al definido en el modelo 2008"

'Global Const Version = "1.00"
'Global Const FechaModificacion = "01/03/2003"
'Global Const UltimaModificacion = "Inicial" '

'**********************************************************************************************************
Dim fs
Dim freg

Global Ternro As Long
Global FechaDesde As Date
Global FechaHasta As Date
Global NroProc As Integer
Global Separador As String

Global Archivo As String
Global pos As Integer
Global strcmdLine  As String
Global Legajo As Long
Global rs As New ADODB.Recordset
Global rsEmpleados As New ADODB.Recordset
Global rsAcumulados As New ADODB.Recordset
Global TotalHoras As Single
Global Fecha As Date
Global X
Global th1 As Integer
Global th2 As Integer
Global strProporcion As String
Global strCantidad As String
Global Desglose As Single
Global empaque As Integer

Global Prueba As String

Global NroProceso As Long


Global pos1 As Byte
Global pos2 As Byte

Sub Main()
Dim PID As String
Dim rs_batch_proceso As New ADODB.Recordset
Dim rs_His_batch_proceso As New ADODB.Recordset
Dim ArrParametros
Dim Nombre_Arch As String
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
    OpenConnection strconexion, objConn
    
        Nombre_Arch = PathFLog & "DesgloseAD" & "-" & NroProceso & ".log"
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Nombre_Arch, True)

    ' Obtengo el Process ID
    PID = GetCurrentProcessId
    Flog.writeline "-----------------------------------------------------------------"
    Flog.writeline "Version = " & Version
    Flog.writeline "Modificacion = " & UltimaModificacion
    Flog.writeline "Fecha = " & FechaModificacion
    Flog.writeline "-----------------------------------------------------------------"
    
    Flog.writeline "PID = " & PID
    Flog.writeline "Inicio Proceso Exportación: " & Now
    Flog.writeline "Cambio el estado del proceso a Procesando"
    'Cambio el estado del proceso a Procesando
    StrSql = "UPDATE batch_proceso SET bprcpid = " & PID & ", bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcestado = 'Procesando', bprcprogreso = 0 WHERE bpronro = " & NroProceso
    objConn.Execute StrSql, , adExecuteNoRecords
    
    Set objFechasHoras.Conexion = objConn
     Flog.writeline "Obtengo los datos del proceso"
    StrSql = " SELECT batch_proceso.bpronro,batch_proceso.bprcfecdesde,batch_proceso.bprcfechasta,batch_procacum.gpanro FROM batch_proceso " & _
             " INNER JOIN batch_procacum ON batch_procacum.bpronro = batch_proceso.bpronro " & _
             " WHERE batch_proceso.bpronro = " & NroProceso
             
    OpenRecordset StrSql, objRs
    
    If Not objRs.EOF Then
        FechaDesde = objRs!bprcfecdesde
        FechaHasta = objRs!bprcfechasta
        NroProc = objRs!gpanro
    End If
    
    Call Exportar_Desglose
            
    StrSql = "DELETE FROM Batch_Procacum WHERE bpronro = " & NroProceso
    objConn.Execute StrSql, , adExecuteNoRecords
    
    StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado' WHERE bpronro = " & NroProceso
    objConn.Execute StrSql, , adExecuteNoRecords
    
    ' -----------------------------------------------------------------------------------
    'FGZ - 22/09/2003
    'Copio el proceso en el historico de batch_proceso y lo borro de batch_proceso
        StrSql = "SELECT * FROM batch_proceso WHERE bpronro =" & NroProceso
        OpenRecordset StrSql, rs_batch_proceso

        
        StrSql = "INSERT INTO His_Batch_Proceso (bpronro,btprcnro,bprcfecha,iduser"
        StrSqlDatos = rs_batch_proceso!bpronro & "," & rs_batch_proceso!btprcnro & "," & _
                 ConvFecha(rs_batch_proceso!bprcfecha) & ",'" & rs_batch_proceso!IdUser & "'"
        
        If Not IsNull(rs_batch_proceso!bprchora) Then
            StrSql = StrSql & ",bprchora"
            StrSqlDatos = StrSqlDatos & ",'" & rs_batch_proceso!bprchora & "'"
        End If
        If Not IsNull(rs_batch_proceso!bprcempleados) Then
            StrSql = StrSql & ",bprcempleados"
            StrSqlDatos = StrSqlDatos & ",'" & rs_batch_proceso!bprcempleados & "'"
        End If
        If Not IsNull(rs_batch_proceso!bprcfecdesde) Then
            StrSql = StrSql & ",bprcfecdesde"
            StrSqlDatos = StrSqlDatos & "," & ConvFecha(rs_batch_proceso!bprcfecdesde)
        End If
        If Not IsNull(rs_batch_proceso!bprcfechasta) Then
            StrSql = StrSql & ",bprcfechasta"
            StrSqlDatos = StrSqlDatos & "," & ConvFecha(rs_batch_proceso!bprcfechasta)
        End If
        If Not IsNull(rs_batch_proceso!bprcestado) Then
            StrSql = StrSql & ",bprcestado"
            StrSqlDatos = StrSqlDatos & ",'" & rs_batch_proceso!bprcestado & "'"
        End If
        If Not IsNull(rs_batch_proceso!bprcparam) Then
            StrSql = StrSql & ",bprcparam"
            StrSqlDatos = StrSqlDatos & ",'" & rs_batch_proceso!bprcparam & "'"
        End If
        If Not IsNull(rs_batch_proceso!bprcprogreso) Then
            StrSql = StrSql & ",bprcprogreso"
            StrSqlDatos = StrSqlDatos & "," & rs_batch_proceso!bprcprogreso
        End If
        If Not IsNull(rs_batch_proceso!bprcfecfin) Then
            StrSql = StrSql & ",bprcfecfin"
            StrSqlDatos = StrSqlDatos & "," & ConvFecha(rs_batch_proceso!bprcfecfin)
        End If
        If Not IsNull(rs_batch_proceso!bprchorafin) Then
            StrSql = StrSql & ",bprchorafin"
            StrSqlDatos = StrSqlDatos & ",'" & rs_batch_proceso!bprchorafin & "'"
        End If
        If Not IsNull(rs_batch_proceso!bprctiempo) Then
            StrSql = StrSql & ",bprctiempo"
            StrSqlDatos = StrSqlDatos & ",'" & rs_batch_proceso!bprctiempo & "'"
        End If
        If Not IsNull(rs_batch_proceso!empnro) Then
            StrSql = StrSql & ",empnro"
            StrSqlDatos = StrSqlDatos & "," & rs_batch_proceso!empnro
        End If
        If Not IsNull(rs_batch_proceso!bprcPid) Then
            StrSql = StrSql & ",bprcPid"
            StrSqlDatos = StrSqlDatos & "," & rs_batch_proceso!bprcPid
        End If
        If Not IsNull(rs_batch_proceso!bprcfecInicioEj) Then
            StrSql = StrSql & ",bprcfecInicioEj"
            StrSqlDatos = StrSqlDatos & "," & ConvFecha(rs_batch_proceso!bprcfecInicioEj)
        End If
        If Not IsNull(rs_batch_proceso!bprcfecFinEj) Then
            StrSql = StrSql & ",bprcfecFinEj"
            StrSqlDatos = StrSqlDatos & "," & ConvFecha(rs_batch_proceso!bprcfecFinEj)
        End If
        If Not IsNull(rs_batch_proceso!bprcUrgente) Then
            StrSql = StrSql & ",bprcUrgente"
            StrSqlDatos = StrSqlDatos & "," & rs_batch_proceso!bprcUrgente
        End If
        If Not IsNull(rs_batch_proceso!bprcHoraInicioEj) Then
            StrSql = StrSql & ",bprcHoraInicioEj"
            StrSqlDatos = StrSqlDatos & ",'" & rs_batch_proceso!bprcHoraInicioEj & "'"
        End If
        If Not IsNull(rs_batch_proceso!bprcHoraFinEj) Then
            StrSql = StrSql & ",bprcHoraFinEj"
            StrSqlDatos = StrSqlDatos & ",'" & rs_batch_proceso!bprcHoraFinEj & "'"
        End If

        StrSql = StrSql & ") VALUES (" & StrSqlDatos & ")"
        objConn.Execute StrSql, , adExecuteNoRecords
        
        
        'Reviso que haya copiado
        StrSql = "SELECT * FROM His_batch_proceso WHERE bpronro =" & NroProceso
        OpenRecordset StrSql, rs_His_batch_proceso
        
        If Not rs_His_batch_proceso.EOF Then
            ' Borro de Batch_proceso
            StrSql = "DELETE FROM Batch_Proceso WHERE bpronro = " & NroProceso
            objConn.Execute StrSql, , adExecuteNoRecords
        End If
    
        If rs_batch_proceso.State = adStateOpen Then rs_batch_proceso.Close
        If rs_His_batch_proceso.State = adStateOpen Then rs_His_batch_proceso.Close
    ' FGZ - 22/09/2003
    ' -----------------------------------------------------------------------------------
           Flog.writeline "Proceso Finalizado Correctamente"
    
    If rs_batch_proceso.State = adStateOpen Then rs_batch_proceso.Close
    If rs_His_batch_proceso.State = adStateOpen Then rs_His_batch_proceso.Close
    If objConn.State = adStateOpen Then objConn.Close
    
    Exit Sub
CE:
    MyRollbackTrans
    If objConn.State = adStateOpen Then objConn.Close
    
End Sub


Public Sub Exportar_Desglose()
Dim Progreso As Single
Dim CEmpleadosAProc As Integer
Dim IncPorc As Single
Dim nroModelo As Long
Dim rs_Modelo As New ADODB.Recordset
Dim Directorio As String
Dim Carpeta
Dim Existe As Boolean
Dim ArchExp
Dim FS1

'' Obtengo los parámetros
'strcmdLine = Command()
'
'If strcmdLine = "" Then Exit Sub
'
'pos1 = 1
'pos2 = InStr(pos1, strcmdLine, ",")
'FechaDesde = CDate(Mid(strcmdLine, pos1, pos2 - pos1))
'
'pos1 = pos2 + 1
'pos2 = InStr(pos1 + 1, strcmdLine, ",")
'FechaHasta = CDate(Mid(strcmdLine, pos1, pos2 - pos1))
'
'pos1 = pos2 + 1
'pos2 = Len(strcmdLine) + 1
'NroProc = Mid(strcmdLine, pos1, pos2 - pos1)

Fecha = FechaDesde

Separador = vbTab
nroModelo = 2008

  'Directorio de exportacion
    StrSql = "SELECT sis_dirsalidas FROM sistema WHERE sisnro = 1 "
    OpenRecordset StrSql, rs
    If Not rs.EOF Then
        Directorio = Trim(rs!sis_dirsalidas)
    End If
     
    StrSql = "SELECT * FROM modelo WHERE modnro = " & nroModelo
    OpenRecordset StrSql, rs_Modelo
    If Not rs_Modelo.EOF Then
        If Not IsNull(rs_Modelo!modarchdefault) Then
            Directorio = Directorio & Trim(rs_Modelo!modarchdefault)
        Else
            Flog.writeline Espacios(Tabulador * 1) & "El modelo no tiene configurada la carpeta desteino. El archivo será generado en el directorio default"
        End If
    Else
        Flog.writeline Espacios(Tabulador * 1) & "No se encontró el modelo " & nroModelo & ". El archivo será generado en el directorio default"
    End If
    Flog.writeline
    
' ------ fgz
' Creo el archivo de texto del desglose
'Archivo = PathFLog & "DesglAD " & NroProceso & ".txt"

'Set fs = CreateObject("Scripting.FileSystemObject")
'Set freg = fs.CreateTextFile(Archivo, True)

' ------ fgz


   Existe = True
    Archivo = Directorio & "\" & "DesglAD " & NroProceso & ".txt"
    Do While Existe
        If fs.FileExists(Archivo) Then
            Archivo = Directorio & "\" & "DesglAD " & NroProceso & ".txt"
        Else
            Existe = False
        End If
    Loop
    Flog.writeline Espacios(Tabulador * 1) & "Se crea el archivo: " & Archivo
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    On Error Resume Next
    Set freg = fs.CreateTextFile(Archivo, True)
    If Err.Number <> 0 Then
        Flog.writeline Espacios(Tabulador * 1) & "La carpeta Destino no existe. Se creará."
        
        Err.Number = 0
        Set FS1 = CreateObject("Scripting.FileSystemObject")
        Set Carpeta = FS1.CreateFolder(Directorio)
        Set ArchExp = fs.CreateTextFile(Archivo, True)
        If Err.Number <> 0 Then
            Flog.writeline Espacios(Tabulador * 1) & "No se pudo crear el archivo"
            Flog.writeline Espacios(Tabulador * 1) & Err.Description
            Exit Sub
        End If
    End If


'OpenConnection strconexion, objConn

' De confrep selecciono los tipos de horas a exportar
' Hora Producción
StrSql = "SELECT confval FROM confrep WHERE confnrocol = 4" & _
" AND repnro = 53"
OpenRecordset StrSql, rs
th1 = rs!confval

' jornada de producción (se mide en fracciones de días)
StrSql = "SELECT confval FROM confrep WHERE confnrocol = 5" & _
" AND repnro = 53"
OpenRecordset StrSql, rs
th2 = rs!confval

StrSql = "SELECT empleado.ternro, empleg FROM empleado "
StrSql = StrSql & "INNER JOIN gti_cab ON gti_cab.ternro = empleado.ternro "
StrSql = StrSql & "WHERE gpanro = " & NroProc
OpenRecordset StrSql, rsEmpleados



' Seteo el incremento de progreso
CEmpleadosAProc = rsEmpleados.RecordCount
If CEmpleadosAProc > 0 Then
    IncPorc = (100 / CEmpleadosAProc)
Else
    IncPorc = 100
End If
Progreso = 0

' Busco el desglose para todos los empleados
Do While Not rsEmpleados.EOF

    Ternro = rsEmpleados!Ternro
    Legajo = rsEmpleados!empleg
    
    Fecha = FechaDesde
    
    'Para el empleado dentro del rango de fechas especificado
    Do While Fecha <= FechaHasta

        StrSql = " SELECT * from gti_achdiario" & _
        " WHERE gti_achdiario.achdfecha = " & ConvFecha(Fecha) & _
        " AND ternro = " & Ternro & " AND thnro = " & th2
        OpenRecordset StrSql, objRs
        
        'Por cada desglose de jornada escribo una línea en el archivo de exportación
        Do While Not objRs.EOF
                
            'Desglose = Round(objrs!achdcanthoras / TotalHoras, 2)
                
            'strProporcion = IIf(Desglose < 1, "0", "") & Trim(Replace(Str(Desglose), ".", ","))
            strCantidad = IIf(objRs!achdcanthoras < 1, "0", "") & Trim(Replace(Str(objRs!achdcanthoras), ".", ","))
            
            freg.Write Legajo & Separador & objRs!achdfecha & Separador & _
            objRs!thnro & Separador & strCantidad & Separador
                            
            'Para cada estructura desglosada
            StrSql = "SELECT * from gti_achdiario_estr" & _
            " INNER JOIN tipoestructura ON tipoestructura.tenro = gti_achdiario_estr.tenro" & _
            " INNER JOIN estructura ON estructura.estrnro = gti_achdiario_estr.estrnro" & _
            " WHERE achdnro = " & objRs!achdnro
            OpenRecordset StrSql, rs
            Do While Not rs.EOF
                freg.Write Chr(34) & rs!estrcodext & Chr(34) & Separador & _
                Chr(34) & rs!estrdabr & Chr(34) & Separador
                rs.MoveNext
            Loop
            
            'Busco el empaque del empleado
            'StrSql = " SELECT estructura.estrnro,estrcodext,estrdabr FROM estruc_actual"
            'StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = estruc_actual.estrnro"
            'StrSql = StrSql & " WHERE estruc_actual.tenro = 36 "
            'StrSql = StrSql & " AND estruc_actual.ternro = " & Ternro
            
            ' 04/08/2003
            'reemplazo estruc_actual por his_estructura
            StrSql = " SELECT estructura.estrnro,estrcodext,estrdabr, his_estructura.htetdesde FROM his_estructura"
            StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro"
            StrSql = StrSql & " WHERE his_estructura.tenro = 36 "
            StrSql = StrSql & " AND his_estructura.htethasta IS NULL "
            StrSql = StrSql & " AND his_estructura.ternro = " & Ternro
            
            OpenRecordset StrSql, rs
            ' ----- 06/05/2003 -----------
            'AIB, FGZ - Hay que controlar que este asignado el empaque
            'empaque = 0 & rs!estrnro
            If Not rs.EOF Then
                freg.Write Chr(34) & rs!estrcodext & Chr(34) & Separador & _
                Chr(34) & rs!estrdabr & Chr(34) & Separador
            End If
            ' ----- 06/05/2003 -----------
            
            'Salto de línea en el .txt
            freg.writeline
            
            'Siguiente desglose
            objRs.MoveNext
        Loop
            
        Fecha = Fecha + 1
    Loop
    
    ' Actualizo el progreso
    Progreso = Progreso + IncPorc
    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & CInt(Progreso) & " WHERE bpronro = " & NroProceso
    objConn.Execute StrSql, , adExecuteNoRecords
    
    rsEmpleados.MoveNext
Loop

freg.Close


End Sub
