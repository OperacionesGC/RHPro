Attribute VB_Name = "mdlExportarDesglAP"
Option Explicit
Dim fs
Dim freg
Global FLOG

Dim Ternro As Long
Dim FechaDesde As Date
Dim FechaHasta As Date
Dim Separador As String

Public Declare Function GetCurrentProcessId Lib "kernel32" () As Long

Sub Main()
Dim Archivo As String
Dim pos As Integer
Dim strcmdLine  As String
Dim Legajo As Long
Dim rs As New ADODB.Recordset
Dim rsCabecera As New ADODB.Recordset
Dim TotalHoras As Single
Dim Proceso As Integer
Dim Cabecera As Integer
Dim PID As String

' carga las configuraciones basicas, formato de fecha, string de conexion,
' tipo de BD y ubicacion del archivo de log
Call CargarConfiguracionesBasicas

strcmdLine = Command()

If IsNumeric(strcmdLine) Then
    Proceso = strcmdLine
Else
    Exit Sub
End If

OpenConnection strconexion, objConn

' Obtengo el Process ID
PID = GetCurrentProcessId
FLOG.writeline "PID = " & PID
'Cambio el estado del proceso a Procesando
StrSql = "UPDATE batch_proceso SET bprcpid = " & PID & ", bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcestado = 'Procesando', bprcprogreso = 0 WHERE bpronro = " & Proceso
objConn.Execute StrSql, , adExecuteNoRecords

Separador = vbTab

StrSql = " SELECT gpadesde, gpahasta, gti_cab.gpanro, cgtinro, ternro from gti_procacum" & _
        " INNER JOIN gti_cab ON gti_cab.gpanro = gti_procacum.gpanro" & _
        " WHERE gti_procacum.gpanro = " & Proceso
OpenRecordset StrSql, rsCabecera

FechaDesde = rsCabecera!gpadesde
FechaHasta = rsCabecera!gpahasta
Cabecera = rsCabecera!cgtinro
    
Archivo = PathFLog & "DesglAP " & Format(FechaDesde, "DD-MM-YYYY") & " al " & Format(FechaHasta, "DD-MM-YYYY") & ".txt"

Set fs = CreateObject("Scripting.FileSystemObject")
Set freg = fs.CreateTextFile(Archivo, True)




Do While Not rsCabecera.EOF

    Cabecera = rsCabecera!cgtinro

    StrSql = "SELECT empleg FROM empleado WHERE ternro = " & rsCabecera!Ternro
    OpenRecordset StrSql, rs
    Legajo = rs!empleg
    
    StrSql = "select sum(achpcanthoras) as horas from gti_achparcial " & _
    " WHERE cgtinro = " & Cabecera
    OpenRecordset StrSql, rs
    
    TotalHoras = 0 & rs!horas
    
    StrSql = "SELECT gti_achparcial.achpnro,achpcanthoras,thnro" & _
    " FROM gti_achparcial" & _
    " WHERE cgtinro = " & Cabecera
    OpenRecordset StrSql, objRs
        
    Do While Not objRs.EOF
        
        freg.Write Legajo & Separador & FechaDesde & Separador & _
        FechaHasta & Separador & _
        objRs!thnro & Separador & objRs!achpcanthoras & Separador & _
        Round(objRs!achpcanthoras / TotalHoras, 2) & Separador
        
        StrSql = "SELECT * FROM gti_achparc_estr" & _
        " INNER JOIN tipoestructura ON tipoestructura.tenro = gti_achparc_estr.tenro" & _
        " INNER JOIN estructura ON estructura.estrnro = gti_achparc_estr.estrnro" & _
        " WHERE achpnro = " & objRs!achpnro
        OpenRecordset StrSql, rs
        
        Do While Not rs.EOF
            freg.Write rs!tenro & Separador & rs!tedabr & Separador & _
            rs!estrnro & Separador & rs!estrdabr & Separador
            
            rs.MoveNext
        Loop
        
        freg.writeline
        objRs.MoveNext
        
    Loop
    rsCabecera.MoveNext
Loop

freg.Close


End Sub
