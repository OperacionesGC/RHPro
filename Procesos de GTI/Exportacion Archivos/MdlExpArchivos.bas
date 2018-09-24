Attribute VB_Name = "MdlExpArchivos"
Option Explicit

Const Version = "1.00"
Const FechaVersion = "28/08/2015" 'LED - CAS-32723 - Salto Grande - GTI Importacion CS-TIME RHPro-Spec - Version Inicial

'---------------------------------------------------------------------------------------------------------------------------------------------
Dim dirsalidas As String
Dim usuario As String
Global Incompleto As Boolean

'-------------------------------------------------------------------------------------------------
'Variables comunes a todas las exportaciones
'-------------------------------------------------------------------------------------------------
Global separador As String
Global directorio As String
Global SeparadorDecimal As String
Global usaencabezado As Boolean
Global DescripcionModelo As String


Public Sub Main()

Dim objconnMain As New ADODB.Connection
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

    Nombre_Arch = PathFLog & "Export_archivo_" & NroProceso & ".log"
    
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
    If Err.Number <> 0 Or Error_Encrypt Then
        Flog.writeline "Problemas en la conexion"
        Exit Sub
    End If

    
    OpenConnection strconexion, objConnProgreso
    If Err.Number <> 0 Or Error_Encrypt Then
        Flog.writeline "Problemas en la conexion"
        Exit Sub
    End If
    On Error GoTo 0
    
    On Error GoTo ME_Main
    
    
    'Cambio el estado del proceso a Procesando
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcprogreso = 0 ,bprcestado = 'Procesando', bprcpid = " & PID & " WHERE bpronro = " & NroProceso
    objConnProgreso.Execute StrSql, , adExecuteNoRecords
    Flog.writeline "Acutaliza el estado en procesando"
    
    'Obtengo los datos del proceso
    StrSql = "SELECT bprcparam, iduser FROM batch_proceso WHERE btprcnro = 457 AND bpronro =" & NroProceso
    OpenRecordset StrSql, rs_batch_proceso
    
    TiempoInicialProceso = GetTickCount
    
    If Not rs_batch_proceso.EOF Then
        bprcparam = IIf(EsNulo(rs_batch_proceso!bprcparam), "", rs_batch_proceso!bprcparam)
        usuario = rs_batch_proceso!iduser
        rs_batch_proceso.Close
        Set rs_batch_proceso = Nothing
        Call proceso_gral(NroProceso, bprcparam)
    Else
        Flog.writeline "no se encontró el proceso"
    End If
    
    TiempoFinalProceso = GetTickCount
    Flog.writeline
    Flog.writeline "--------------------------------------------------------------"
    Flog.writeline "Tiempo del proceso (milisegundos): " & (TiempoFinalProceso - TiempoInicialProceso)
    
    If Incompleto Then
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Incompleto' WHERE bpronro = " & NroProceso
        objConnProgreso.Execute StrSql, , adExecuteNoRecords
    Else
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado' WHERE bpronro = " & NroProceso
        objConnProgreso.Execute StrSql, , adExecuteNoRecords
    End If
    
Fin:
    Flog.Close
    If objConnProgreso.State = adStateOpen Then objConnProgreso.Close
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
        StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcprogreso = 0, bprcestado = 'Error', bprcpid = " & PID & " WHERE bpronro = " & NroProceso
        objConnProgreso.Execute StrSql, , adExecuteNoRecords
    MyCommitTrans
End Sub

Public Sub proceso_gral(ByVal bpronro As Long, ByVal parametros As String)
' Parametros:  1º Numero de modelo
'              2º si toma el directorio de la configuracion del modelo (-1) o si de lo configurado en sis_expseguridad de la tabla sistema
Dim arrayParametros
Dim rs_datos As New ADODB.Recordset

    arrayParametros = Split(parametros, "@")
    
    'busco el directorio donde se grabaran los archivos
    StrSql = "SELECT sis_expseguridad, sis_dirsalidas FROM sistema WHERE sisnro = 1 "
    OpenRecordset StrSql, rs_datos
    If Not rs_datos.EOF Then
        If CLng(arrayParametros(1)) = -1 Then
            directorio = Trim(rs_datos!sis_dirsalidas)
        Else
            directorio = Trim(rs_datos!sis_expseguridad)
        End If
       
    Else
       Flog.writeline "No esta configurado el directorio de sistema."
       Exit Sub
    End If
    
    StrSql = "SELECT * FROM modelo WHERE modnro = " & arrayParametros(0)
    OpenRecordset StrSql, rs_datos
    If Not rs_datos.EOF Then
        If CLng(arrayParametros(1)) = -1 Then
            directorio = directorio & Trim(rs_datos!modarchdefault)
        End If
        
        separador = IIf(Not IsNull(rs_datos!modseparador), rs_datos!modseparador, ",")
        SeparadorDecimal = IIf(Not IsNull(rs_datos!modsepdec), rs_datos!modsepdec, ".")
        usaencabezado = IIf(Not IsNull(rs_datos!modencab), CBool(rs_datos!modencab), False)
        DescripcionModelo = rs_datos!moddesc
        
        Flog.writeline Espacios(Tabulador * 1) & "Modelo " & arrayParametros(0) & " " & DescripcionModelo
        Flog.writeline
        Flog.writeline Espacios(Tabulador * 1) & "Directorio de exportacion " & directorio
     Else
        Flog.writeline Espacios(Tabulador * 1) & "No se encontró el modelo " & arrayParametros(0)
        Exit Sub
    End If
    
    
    'busco el modelo a ajecutar
    Select Case CLng(arrayParametros(0))
        Case 2006: 'CAS-32723 - Salto Grande - GTI Importacion CS-TIME RHPro-Spec - Exportacion de tarjetas de empleado
            Call modelo_2006(arrayParametros(0), arrayParametros(2))
    End Select
    
    

    
End Sub

