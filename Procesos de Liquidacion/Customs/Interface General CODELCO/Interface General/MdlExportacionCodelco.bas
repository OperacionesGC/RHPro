Attribute VB_Name = "MdlExportar"
Option Explicit

Global NroProceso As Long

Global f
'Global HuboError As Boolean
Global Path
Global NArchivo
Global freg
Global NroLinea As Long
Global LineaCarga As Long
Global directorio As String


Global Separador As String
Global SeparadorDecimal As String
Global UsaEncabezado As Boolean

Global ErrCarga
Global LineaError
Global LineaOK

Global NombreArchivo As String
Global NroModelo As Long
Global TipoExport As Integer


Global usuario As String

Global FechaIni As String
Global FechaFin As String



Public Sub Main()
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento inicial de Interface.
' Autor      : FGZ
' Fecha      : 29/07/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim objconnMain As New ADODB.Connection
Dim strcmdLine
Dim Nombre_Arch As String
Dim Nombre_Arch_Errores  As String
Dim rs_batch_proceso As New ADODB.Recordset
Dim bprcparam As String
Dim PID As String

' carga las configuraciones basicas, formato de fecha, string de conexion,
' tipo de BD y ubicacion del archivo de log

Call CargarConfiguracionesBasicas
    
'Abro la conexion
    OpenConnection strconexion, objConn
    OpenConnection strconexion, objconnProgreso
    
    strcmdLine = Command()
    If IsNumeric(strcmdLine) Then
        NroProcesoBatch = strcmdLine
    Else
        Exit Sub
    End If
    
    Nombre_Arch = PathFLog & "Migracion_Interface" & "-" & NroProcesoBatch & ".log"
    Nombre_Arch_Errores = PathFLog & "Migracion_Interface_Errores" & "-" & NroProcesoBatch & ".log"
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Nombre_Arch, True)
    Set FlogE = fs.CreateTextFile(Nombre_Arch_Errores, True)
    
    ' Obtengo el Process ID
    PID = GetCurrentProcessId
    Flog.writeline Espacios(Tabulador * 0) & "PID = " & PID
    
    'Cambio el estado del proceso a Procesando
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcestado = 'Procesando', bprcprogreso = 0, bprcpid = " & PID & " WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    'Obtengo los datos del proceso
    'StrSql = "SELECT * FROM batch_proceso WHERE btprcnro = 65 AND bpronro =" & NroProcesoBatch
    StrSql = "SELECT * FROM batch_proceso WHERE btprcnro = 70 AND bpronro =" & NroProcesoBatch
    OpenRecordset StrSql, rs_batch_proceso
    
    
    If Not rs_batch_proceso.EOF Then
        bprcparam = rs_batch_proceso!bprcparam
        usuario = rs_batch_proceso!iduser
        rs_batch_proceso.Close
        Set rs_batch_proceso = Nothing
        Call LevantarParamteros(bprcparam)
        LineaCarga = 0
        Flog.writeline "LLAMO GENERARARCHIVO"
        Call GenerarArchivo(NombreArchivo, TipoExport)
        Flog.writeline "SALIO DE GENERARARCHIVO"
    End If
    
    If Not HuboError Then
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcprogreso = 100, bprcestado = 'Procesado' WHERE bpronro = " & NroProcesoBatch
    Else
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcprogreso = 100, bprcestado = 'Error' WHERE bpronro = " & NroProcesoBatch
    End If
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    objConn.Close
    objconnProgreso.Close
    Flog.Close
    FlogE.Close
End Sub


Private Sub GenerarArchivo(ByVal NombreArchivo As String, ByVal TipoExport As Integer)
' --------------------------------------------------------------
' Descripcion: Exportación de Novedades de Liq.
' Autor: ?
' Ultima modificacion: FGZ - 29/07/2003
' --------------------------------------------------------------

Dim strcmdLine As String

Dim Archivo As String
Dim pos As Integer

Dim NAmbito As String
Dim Linea As String

Dim objRs As New ADODB.Recordset
Flog.writeline "ENTRO EN GENERARARCHIVO"

    StrSql = "SELECT sis_direntradas FROM sistema WHERE sisnro = 1 "

Flog.writeline StrSql
    
    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then
        directorio = Trim(objRs!sis_direntradas)
    Else
        Flog.writeline Espacios(Tabulador * 1) & "No se encontró el registro de la tabla sistema nro 1"
        Exit Sub
    End If

    NombreArchivo = directorio & "\ExportacionCodelco\" & NombreArchivo

Flog.writeline "NombreArchivo = " & NombreArchivo

    Set fs = CreateObject("Scripting.FileSystemObject")
    Set freg = fs.CreateTextFile(NombreArchivo, True)

    Select Case TipoExport
    
        Case 0 ' Tabla Formacion
            Flog.writeline "LLAMO EvaFormacion"
            Call EvaFormacion
        
        Case 1 ' Tabla Cabecera
            Flog.writeline "LLAMO EvaCabeceras"
            Call EvaCabeceras
        
        Case 2 ' Tabla Secciones
            Flog.writeline "LLAMO EvaSecciones"
            Call EvaSecciones
        
                    
    End Select

    Exit Sub

CE:
    Debug.Print StrSql
    Debug.Print Err.Description

    HuboError = True
    
    MyRollbackTrans
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 0) & "**********************************************************"
    Flog.writeline Espacios(Tabulador * 0) & "Error. " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    Flog.writeline Espacios(Tabulador * 0) & "Error: " & Err.Number
    Flog.writeline Espacios(Tabulador * 0) & "Decripcion: " & Err.Description
    Flog.writeline
    If InStr(1, Err.Description, "ODBC") > 0 Then
        'Fue error de Consulta de SQL
        Flog.writeline
        Flog.writeline Espacios(Tabulador * 0) & "SQL Ejecutado: " & StrSql
        Flog.writeline
    End If
    Flog.writeline Espacios(Tabulador * 0) & "**********************************************************"
    Flog.writeline


End Sub

Public Sub EvaCabeceras()

Dim rs As New ADODB.Recordset
Dim rs_pun As New ADODB.Recordset
Dim NAmbito As String
Flog.writeline "ENTRO EvaCabeceras"
    StrSql = " SELECT empleado.empleg, evacab.evacabnro, evaevento.evaevefdesde, "
    StrSql = StrSql & " evaevento.evaevefhasta, evacab.puntajemanual "
    StrSql = StrSql & " FROM evaevento "
    StrSql = StrSql & " INNER JOIN evacab ON evacab.evaevenro = evaevento.evaevenro "
    StrSql = StrSql & " INNER JOIN empleado ON empleado.ternro = evacab.empleado "
    StrSql = StrSql & " WHERE evaevento.evaevefdesde >= " & FechaIni
    StrSql = StrSql & " AND evaevento.evaevefhasta <= " & FechaFin
Flog.writeline StrSql
    OpenRecordset StrSql, rs
    
   
    Do While Not rs.EOF
        
        ' chr(34) Son las Comillas
        
        StrSql = " SELECT * FROM evapuntaje "
        StrSql = StrSql & " INNER JOIN evatipoobj ON evatipoobj.evatipobjnro = evapuntaje.evatipobjnro "
        StrSql = StrSql & " WHERE evapuntaje.evacabnro = " & rs!evacabnro
        
        OpenRecordset StrSql, rs_pun
        
        NAmbito = ""
        
        Do While Not rs_pun.EOF
        
            If NAmbito = "" Then
                NAmbito = rs_pun!puntaje
            Else
                NAmbito = NAmbito & ";" & rs_pun!puntaje
            End If
            
            rs_pun.MoveNext
        
        Loop
            
        freg.writeline rs!empleg & ";" & _
                       rs!evacabnro & ";" & _
                       Format(rs!evaevefdesde, "DD/MM/YYYY") & ";" & _
                       Format(rs!evaevefhasta, "DD/MM/YYYY") & ";" & _
                       NAmbito & ";" & _
                       rs!puntajemanual

        rs.MoveNext
        
    Loop
    
Fin:
    If rs.State = adStateOpen Then rs.Close
    Set rs = Nothing
    
    If rs_pun.State = adStateOpen Then rs_pun.Close
    Set rs_pun = Nothing
    
    Exit Sub
    
End Sub


Public Sub EvaSecciones()

Dim rs As New ADODB.Recordset
Dim rs_pun As New ADODB.Recordset
Dim NAmbito As String

Flog.writeline "ENTRO EvaSecciones"
    StrSql = " SELECT evacab.evacabnro, evasecc.titulo, evadetevldor.evldorcargada, "
    StrSql = StrSql & " evadetevldor.fechacar "
    StrSql = StrSql & " FROM evaevento "
    StrSql = StrSql & " INNER JOIN evacab ON evacab.evaevenro = evaevento.evaevenro "
    StrSql = StrSql & " INNER JOIN evadetevldor ON evadetevldor.evacabnro = evacab.evacabnro "
    StrSql = StrSql & " INNER JOIN evasecc ON evasecc.evaseccnro =  evadetevldor.evaseccnro "
    StrSql = StrSql & " WHERE evaevento.evaevefdesde >= " & FechaIni
    StrSql = StrSql & " AND evaevento.evaevefhasta <= " & FechaFin
Flog.writeline StrSql
    OpenRecordset StrSql, rs
    
    Do While Not rs.EOF
        
        ' chr(34) Son las Comillas
        
        freg.writeline rs!evacabnro & ";" & _
                       rs!titulo & ";" & _
                       rs!evldorcargada & ";" & _
                       Format(rs!fechacar, "dd/mm/yyyy")

        rs.MoveNext
        
    Loop
    
Fin:
    If rs.State = adStateOpen Then rs.Close
    Set rs = Nothing
    Exit Sub
    
End Sub


Public Sub EvaFormacion()

Flog.writeline "ENTRO EvaFormacion"
Dim rs As New ADODB.Recordset
Dim rs_pun As New ADODB.Recordset
Dim NAmbito As String
   
StrSql = " SELECT evaevento.evaevedesabr,estructura.estrdabr, empleado.empleg empleado, "
StrSql = StrSql & " empleado.terape, empleado.ternom, emp.empleg evaluador, evanotas.evanotadesc, evaevento.evaevefecha "
StrSql = StrSql & " FROM evaevento "
StrSql = StrSql & " INNER JOIN evacab ON evacab.evaevenro = evaevento.evaevenro"
StrSql = StrSql & " INNER JOIN empleado ON empleado.ternro = evacab.empleado"
StrSql = StrSql & " INNER JOIN his_estructura on his_estructura.ternro = empleado.ternro and his_estructura.tenro=44"
StrSql = StrSql & " INNER JOIN estructura on his_estructura.estrnro = estructura.estrnro "
StrSql = StrSql & " INNER JOIN evadetevldor ON evadetevldor.evacabnro = evadetevldor.evacabnro"
StrSql = StrSql & " INNER JOIN empleado emp ON emp.ternro = evadetevldor.evaluador"
StrSql = StrSql & " INNER JOIN evanotas on evanotas.evldrnro = evadetevldor.evldrnro AND evanotas.evatnnro=4"
StrSql = StrSql & " WHERE evaevento.evaevefdesde >= " & FechaIni & " AND evaevento.evaevefhasta <= " & FechaFin
OpenRecordset StrSql, rs

If rs.EOF Then
    Flog.writeline " No existen empleados con evaluaciones entre las fechas seleccionadas"
    MyRollbackTrans
    Exit Sub
End If
Dim asd
Do While Not rs.EOF
    Flog.writeline "Exportando lista de evaluaciones entre las fechas " & FechaIni & " y " & FechaFin & "."
    freg.writeline rs("evaevedesabr") & ";" & rs("estrdabr") & ";" & rs("empleado") & ";" & rs("ternom") & ";" & rs("terape") & ";" & rs("evaluador") & ";" & rs("evanotadesc") & ";" & rs("evaevefecha") & ";"
    rs.MoveNext
Loop

rs.Close
End Sub


Public Sub LevantarParamteros(ByVal parametros As String)
Dim pos1 As Integer
Dim pos2 As Integer

Dim NombreArchivo1 As String
Dim NombreArchivo2 As String
Dim NombreArchivo3 As String

Dim Fecha As Date

Separador = "@"
If Not IsNull(parametros) Then
    If Len(parametros) >= 1 Then

        'Nro de Modelo
        pos1 = 1
        pos2 = InStr(pos1, parametros, Separador) - 1
        Fecha = Mid(parametros, pos1, pos2 - pos1 + 1)
        
        FechaFin = ConvFecha(Fecha)
        FechaIni = ConvFecha(DateAdd("d", -365, Fecha))
        
        pos1 = pos2 + 2
        pos2 = InStr(pos1, parametros, Separador) - 1
        NombreArchivo = Mid(parametros, pos1, pos2 - pos1 + 1)
        
        'Nombre del archivo a levantar
        pos1 = pos2 + 2
        pos2 = InStr(pos1, parametros, Separador) - 1
        If pos2 > 0 Then
            TipoExport = Mid(parametros, pos1, pos2 - pos1 + 1)
        Else
            pos2 = Len(parametros)
            TipoExport = Mid(parametros, pos1, pos2 - pos1 + 1)
        End If
        
    End If
End If

End Sub






