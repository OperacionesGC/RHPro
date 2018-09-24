Attribute VB_Name = "MdlRHProHelpnex"
Option Explicit

Global Const Version = "1.00"
Global Const FechaModificacion = "22/05/2015"
Global Const UltimaModificacion = "" ' Stremel Sebastian - CAS-27585 - VILLA MARIA - VISTA DE DATOS [Entrega 2] - Cada vez que se inserta un empleado, se dispara el proceso y lo inserta en rhpro_helpnex

Dim usuario As String
Dim Incompleto As Boolean



Public Sub Main()

Dim objconnMain As New ADODB.Connection
Dim strCmdLine
Dim Nombre_Arch As String
Dim rs_batch_proceso As New ADODB.Recordset
Dim PID As String
Dim bprcparam As String
Dim ArrParametros

Dim param
Dim ternroAux As Integer

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

    Nombre_Arch = PathFLog & "RHPro_Helpnex" & NroProcesoBatch & ".log"
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Nombre_Arch, True)
    
    
    ' Obtengo el Process ID
    PID = GetCurrentProcessId
    Flog.writeline "-------------------------------------------------"
    Flog.writeline "Version                  : " & Version
    Flog.writeline "Fecha Ultima Modificacion: " & FechaModificacion
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

    
    OpenConnection strconexion, objconnProgreso
    If Err.Number <> 0 Or Error_Encrypt Then
        Flog.writeline "Problemas en la conexion"
        Exit Sub
    End If
    On Error GoTo 0
    
    On Error GoTo ME_Main
    
    
    'Cambio el estado del proceso a Procesando
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcprogreso = 0 ,bprcestado = 'Procesando', bprcpid = " & PID & " WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    Flog.writeline "Acutaliza el estado en procesando"
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE btprcnro = 448 AND bpronro =" & NroProcesoBatch
    OpenRecordset StrSql, rs_batch_proceso
    
    TiempoInicialProceso = GetTickCount
    
    If Not rs_batch_proceso.EOF Then
        bprcparam = IIf(EsNulo(rs_batch_proceso!bprcparam), "", rs_batch_proceso!bprcparam)
        usuario = rs_batch_proceso!iduser
        rs_batch_proceso.Close
        Set rs_batch_proceso = Nothing
        Call insertarDatos(bprcparam)
    Else
        Flog.writeline "No se encontró el proceso"
    End If
    
    TiempoFinalProceso = GetTickCount
    Flog.writeline
    Flog.writeline "--------------------------------------------------------------"
    Flog.writeline "Tiempo del proceso (milisegundos): " & (TiempoFinalProceso - TiempoInicialProceso)
    
    If HuboError Then
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Error' WHERE bpronro = " & NroProcesoBatch
        objconnProgreso.Execute StrSql, , adExecuteNoRecords
    Else
        If Incompleto Then
            StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Incompleto' WHERE bpronro = " & NroProcesoBatch
            objconnProgreso.Execute StrSql, , adExecuteNoRecords
        Else
            StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado', bprcprogreso=100 WHERE bpronro = " & NroProcesoBatch
            objconnProgreso.Execute StrSql, , adExecuteNoRecords
        End If
    End If
    
    
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

Sub insertarDatos(ByVal param As String)

On Error GoTo Error
Dim StrSql As String
Dim rs_datos As New ADODB.Recordset

Dim datos
Dim ternro As Long
Dim empleg As Long
Dim tipoOp As String
Dim teLugarDeTrabajo As Long
Dim estrnro As String
Dim codigoConexion As Integer
Dim cn_helpnex As New ADODB.Connection


Dim codGrupo As Integer
Dim idPerfil As Integer
Dim rhtexto As String
Dim Encontro As Boolean
Encontro = False
teLugarDeTrabajo = 0

datos = Split(param, "@")
If UBound(datos) <> 1 Then
    Flog.writeline "La cantidad de parametros recibida es incorrecta.Se aborta el proceso"
    Incompleto = True
    Exit Sub
End If

If Not EsNulo(datos(0)) And IsNumeric(datos(0)) Then
    ternro = datos(0)
    '----------------------Busco el legajo del empleado------------------------------
    StrSql = "SELECT empleg FROM empleado"
    StrSql = StrSql & " WHERE ternro=" & ternro
    OpenRecordset StrSql, rs_datos
    If Not rs_datos.EOF Then
        empleg = rs_datos!empleg
        Flog.writeline "Legajo del empleado: " & empleg
    Else
        Flog.writeline "No se encontro el legajo del empleado.Se aborta el proceso."
        Incompleto = True
        Exit Sub
    End If
    rs_datos.Close
    '--------------------------------------------------------------------------------
Else
    Incompleto = True
    Flog.writeline "El parametro 1 es incorrecto.Se aborta el proceso"
End If


If Not EsNulo(datos(1)) Then
    tipoOp = datos(1)
    Flog.writeline "Tipo Operacion:" & tipoOp
End If

'---------------------------LEVANTO LOS DATOS DEL CONFREP---------------------------
Flog.writeline "Levanto los datos del confrep"
StrSql = " SELECT * FROM confrep "
StrSql = StrSql & " WHERE repnro= 465 "
StrSql = StrSql & " AND confnrocol=3"
OpenRecordset StrSql, rs_datos
If Not rs_datos.EOF Then
    If Not EsNulo(rs_datos!confval) Then
        teLugarDeTrabajo = rs_datos!confval
        Flog.writeline "Lugar de trabajo:" & teLugarDeTrabajo
    Else
        Flog.writeline "El tipo de estructura Lugar de Trabajo esta mal configurado.Se aborta el proceso"
        Incompleto = True
        Exit Sub
    End If
Else
    Flog.writeline "El tipo de estructura Lugar de Trabajo no esta configurado.Se aborta el proceso"
    Incompleto = True
    Exit Sub
End If
rs_datos.Close

StrSql = " SELECT estrnro FROM his_estructura "
StrSql = StrSql & " WHERE tenro=" & teLugarDeTrabajo & " AND ternro=" & ternro
StrSql = StrSql & " AND ((htetdesde<=" & ConvFecha(Date) & ") AND (htethasta >=" & ConvFecha(Date) & " or htethasta is null))"
OpenRecordset StrSql, rs_datos
If Not rs_datos.EOF Then
    If Not EsNulo(rs_datos!estrnro) Then
        estrnro = rs_datos!estrnro
    Else
        Flog.writeline "El empleado no tiene la estructura lugar de trabajo.Se aborta el proceso"
        Incompleto = True
        Exit Sub
    End If
Else
    estrnro = 0
    Flog.writeline "El empleado no tiene la estructura lugar de trabajo.Se aborta el proceso"
    Incompleto = True
    Exit Sub
End If
rs_datos.Close


If tipoOp = "A" Then
    '---levanto del confrep el codigo de la conexion para el lugar de trabajo---
    StrSql = " SELECT * FROM confrep "
    StrSql = StrSql & " WHERE repnro=465"
    StrSql = StrSql & " AND upper(conftipo)='CON' "
    Flog.writeline "Query: " & StrSql
    OpenRecordset StrSql, rs_datos
    If Not rs_datos.EOF Then
        Do While Not rs_datos.EOF
            If Not Encontro Then
                If InStr("," & Replace(rs_datos!confval2, " ", "") & ",", "," & estrnro & ",") > 0 Then
                    codigoConexion = rs_datos!confval
                    Encontro = True
                End If
            End If
        rs_datos.MoveNext
        Loop
        Flog.writeline "Codigo de conexion: " & codigoConexion
    Else
        codigoConexion = 0
        Flog.writeline "No se encontro la conexion correspondiente.Se aborta el proceso."
        Incompleto = True
        Exit Sub
    End If
    rs_datos.Close
    '--------------------------------------------------------------------------
    
    '------------------------BUSCO EL STRING DE CONEXION-----------------------
    StrSql = " SELECT cnstring FROM conexion "
    StrSql = StrSql & " WHERE cnnro=" & codigoConexion
    OpenRecordset StrSql, rs_datos
    If rs_datos.EOF Then
        strconexion = ""
        Flog.writeline "No se encontro el string conexion. Se aborta el proceso."
        Incompleto = True
        Exit Sub
    Else
        strconexion = rs_datos!cnstring
    End If
    rs_datos.Close
    '---------------------------------------------------------------------------
    
    'Abro la conexion a Helpnex
    cn_helpnex.ConnectionString = strconexion
    cn_helpnex.Open
    
    'RECUPERO LOS DATOS NECESARIOS A INSERTAR
    StrSql = " SELECT idPerfilUsuario, codigoGrupo, texto "
    StrSql = StrSql & " FROM RHPRO_Relaciones "
    StrSql = StrSql & " WHERE codigoGrupo=1 "
    OpenRecordsetExt StrSql, rs_datos, cn_helpnex
    If Not rs_datos.EOF Then
        codGrupo = rs_datos!CodigoGrupo
        idPerfil = rs_datos!idperfilusuario
        rhtexto = rs_datos!Texto
    Else
        Flog.writeline "No se encontraron los datos del grupo. Se aborta el proceso"
        Incompleto = True
        Exit Sub
    End If
    rs_datos.Close
    
    'INSERTO LOS DATOS
    StrSql = " SELECT * FROM rhpro_helpnex "
    StrSql = StrSql & " WHERE empleg=" & empleg & " AND ternro=" & ternro
    OpenRecordset StrSql, rs_datos
    If rs_datos.EOF Then
        Flog.writeline "Se va a insertar el legajo:" & empleg
        StrSql = "INSERT INTO rhpro_helpnex "
        StrSql = StrSql & " (ternro, empleg, IDPerfilAcceso, CodigoGrupo, rhptexto) "
        StrSql = StrSql & " VALUES "
        StrSql = StrSql & " ( "
        StrSql = StrSql & ternro & ","
        StrSql = StrSql & empleg & ","
        StrSql = StrSql & idPerfil & ","
        StrSql = StrSql & codGrupo & ","
        StrSql = StrSql & "'" & rhtexto & "'"
        StrSql = StrSql & " ) "
        objConn.Execute StrSql, , adExecuteNoRecords
        Flog.writeline "Se inserto el empleado"
    Else
        Flog.writeline "El legajo:" & empleg & "YA EXISTE EN EL SISTEMA, NO SE INSERTA."
    End If
    
End If

Exit Sub

Error:
    Flog.writeline "Se produjo un error: " & Err.Description
    Flog.writeline "Se aborta el proceso."
    HuboError = True
    Exit Sub

End Sub

Public Sub OpenRecordsetExt(strSQLQuery As String, ByRef objRs As ADODB.Recordset, ByVal objConnE As ADODB.Connection, Optional lockType As LockTypeEnum = adLockReadOnly)
' ---------------------------------------------------------------------------------------------
' Descripcion: Abre recordset de conexion externa
' Autor      : Martin Ferraro
' Fecha      : 18/04/2009
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------

Dim pos1 As Long
Dim pos2 As Long
Dim aux As String

    'Abre un recordset con la consulta strSQLQuery
    If objRs.State <> adStateClosed Then
        If objRs.lockType <> adLockReadOnly Then objRs.UpdateBatch
        objRs.Close
    End If
    
    objRs.CacheSize = 500

    objRs.Open strSQLQuery, objConnE, adOpenDynamic, lockType, adCmdText
    
    Cantidad_de_OpenRecordset = Cantidad_de_OpenRecordset + 1
    
End Sub
