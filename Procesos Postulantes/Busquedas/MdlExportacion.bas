Attribute VB_Name = "MdlBusquedas"
Option Explicit

'Version: 1.01
'
'Const Version = 1.02
'Const FechaVersion = "22/09/2005"

'Const Version = 1.03 ' Fapitalle N. - se agrega el token $FECHACV$ para reemplazar en la sql de busqueda
'Const FechaVersion = "06/04/2006"

'Const Version = 1.04 ' Fapitalle N. - se agrega el token $ULTIMOESTADO$ para reemplazar en la sql de busqueda
'Const FechaVersion = "18/05/2006"

'Global Const Version = 1.05
'Global Const FechaVersion = "06/10/2009"   'Encriptacion de string connection
'Global Const UltimaModificacion = "Manuel Lopez"
'Global Const UltimaModificacion1 = "Encriptacion de string connection"

'Global Const Version = 1.06
'Global Const FechaVersion = "12/01/2010"   'Agregado de logs
'Global Const UltimaModificacion = "MB"
'Global Const UltimaModificacion1 = "Agregado de logs"

'Global Const Version = 1.07
'Global Const FechaVersion = "17/05/2011"   'Correccion en los estados
'Global Const UltimaModificacion = "Lisandro Moro"
'Global Const UltimaModificacion1 = "Correccion en los estados"

Global Const Version = 1.08
Global Const FechaVersion = "14/06/2011"
Global Const UltimaModificacion = "Martinez Nicolas & Lisandro Moro"
Global Const UltimaModificacion1 = "Se agrego parametro nuevo que permite determinar si se deben incluir postulantes ya seleccionados en otros procesos de busqueda." & " - Correccion en los estados"

Global IdUser As String
Global Fecha As Date
Global hora As String

Const ULTIMO_ESTADO = "tercero.ternro IN (" & _
                        "SELECT DISTINCT ternro FROM pos_seguimiento ps1 WHERE NOT EXISTS(" & _
                            " SELECT ternro FROM pos_seguimiento ps2" & _
                            " WHERE ps2.actnro=ps1.actnro" & _
                            " AND ps2.ternro=ps1.ternro" & _
                            " AND ps2.estnro<>ps1.estnro" & _
                            " AND ps2.segfec>ps1.segfec)"




Public Sub Main()
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento inicial del Proceso.
' Autor      : FGZ
' Fecha      : 07/09/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim strCmdLine
Dim Nombre_Arch As String
Dim HuboError As Boolean
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
    
    TiempoInicialProceso = GetTickCount
    
    ' carga las configuraciones basicas, formato de fecha, string de conexion,
    ' tipo de BD y ubicacion del archivo de log
    Call CargarConfiguracionesBasicas

    Nombre_Arch = PathFLog & "Busqueda" & "-" & NroProcesoBatch & ".log"
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Nombre_Arch, True)
    
    ' Obtengo el Process ID
    PID = GetCurrentProcessId
    Flog.writeline Espacios(Tabulador * 0) & "-------------------------------------------------"
    Flog.writeline Espacios(Tabulador * 0) & "Version                  : " & Version
    Flog.writeline Espacios(Tabulador * 0) & "Fecha Ultima Modificacion: " & FechaVersion
    Flog.writeline Espacios(Tabulador * 0) & "Modificacion = " & UltimaModificacion1
    Flog.writeline Espacios(Tabulador * 0) & "PID                      : " & PID
    Flog.writeline Espacios(Tabulador * 0) & "-------------------------------------------------"
    Flog.writeline
    
    'Abro la conexion
    On Error Resume Next
    OpenConnection strconexion, objConn
    If Err.Number <> 0 Or Error_Encrypt Then
     Flog.writeline Espacios(Tabulador * 0) & "Problemas en la conexion"
    Exit Sub
    End If
    
    On Error Resume Next
    OpenConnection strconexion, objconnProgreso
    If Err.Number <> 0 Or Error_Encrypt Then
     Flog.writeline Espacios(Tabulador * 0) & "Problemas en la conexion"
    Exit Sub
    End If
    On Error GoTo 0
    
    On Error GoTo ME_Main
    
    'Cambio el estado del proceso a Procesando y el PID
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcprogreso = 0, bprcestado = 'Procesando', bprcpid = " & PID & " WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE btprcnro = 89 AND bpronro =" & NroProcesoBatch
    OpenRecordset StrSql, rs_batch_proceso
    
    TiempoInicialProceso = GetTickCount
    
    If Not rs_batch_proceso.EOF Then
        IdUser = rs_batch_proceso!IdUser
        Fecha = rs_batch_proceso!bprcfecha
        hora = rs_batch_proceso!bprchora
        bprcparam = rs_batch_proceso!bprcparam
        
        rs_batch_proceso.Close
        Set rs_batch_proceso = Nothing
        
        Call LevantarParametros(NroProcesoBatch, bprcparam)
    End If
    
    TiempoFinalProceso = GetTickCount
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 0) & "=================================================="
    Flog.writeline Espacios(Tabulador * 0) & "Tiempo del proceso (milisegundos): " & (TiempoFinalProceso - TiempoInicialProceso)
    Flog.writeline Espacios(Tabulador * 0) & "=================================================="
    
    If Not HuboError Then
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado' WHERE bpronro = " & NroProcesoBatch
    Else
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Error' WHERE bpronro = " & NroProcesoBatch
    End If
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
Fin:
    Flog.Close
    If objConn.State = adStateOpen Then objConn.Close
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
    GoTo Fin:
End Sub

Public Sub Generacion(ByVal busnro As Long, ByVal formal As Boolean, ByVal reqpernro As Long, ByVal agregar As Boolean, ByVal seleccionados As Boolean)
' --------------------------------------------------------------------------------------------
' Descripcion: Procedimiento que genera la busqueda de postulantes
' Autor      : Fernando Favre
' Fecha      : 17/05/2005
' Ult. Mod   :
' Fecha      :
' --------------------------------------------------------------------------------------------
Dim StrSqlw As String
Dim StrSql2 As String

'se definen dos variables nuevas para el caso de buscar la fecha de presentacion
'en dos tablas distintas
Dim StrSqlw_p As String
Dim StrSqlw_e As String

Dim Seguir As Boolean
Dim busempact As Boolean
Dim busempina As Boolean
Dim buspos As Boolean
Dim caracsql As String
Dim reqbusnro As Long
Dim progresoInc As Single

Dim busestposnro As String

'Registros
Dim rs_pos_busqueda As New ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim rs_Pos_Terreqbus As New ADODB.Recordset
Dim rs_eliminar As New ADODB.Recordset

Dim busestposnroArr
Dim a

MyBeginTrans

StrSql = "SELECT estposnro "
'StrSql = "SELECT busempact, busempina, buspos "
StrSql = StrSql & "FROM pos_busqueda "
StrSql = StrSql & "WHERE pos_busqueda.busnro=" & busnro
OpenRecordset StrSql, rs_pos_busqueda
If Not rs_pos_busqueda.EOF Then
    'busempact = rs_pos_busqueda("busempact")
    'busempina = rs_pos_busqueda("busempina")
    'buspos = rs_pos_busqueda("buspos")
    
    busestposnro = rs_pos_busqueda("estposnro")
    
    '17/05/2011 - Lisandor Moro - Valido si busco por activos e inactivos.
    busempact = False
    busempina = False
    
    'StrSql2 = " SELECT estposact "
    'StrSql2 = StrSql2 & " FROM pos_estpos "
    'StrSql2 = StrSql2 & " WHERE estposnro IN (" & busestposnro & ")"
    'OpenRecordset StrSql2, rs
    'If Not rs.EOF Then
    '    Do While Not rs.EOF
    '        If Not IsNull(rs("estposact")) Then
    '            If CLng(rs("estposact")) = -1 Then
    '                busempact = True
    '            End If
    '            If CLng(rs("estposact")) = 0 Then
    '                busempina = True
    '            End If
    '        End If
    '        rs.MoveNext
    '    Loop
    'Else
    '    busempact = False
    '    busempina = False
    'End If
    'rs.Close
    
    busestposnroArr = Split(busestposnro, ",")
    For a = 0 To UBound(busestposnroArr)
        If CLng(busestposnroArr(a)) = 1 Then
            busempact = True
        End If
        If CLng(busestposnroArr(a)) = 2 Then
            busempina = True
        End If
    Next
    
    
    Flog.writeline Espacios(Tabulador * 0) & "Activos: " & busempact
    Flog.writeline Espacios(Tabulador * 0) & "Inactivos: " & busempina
    'If InStr(busestposnro, 1) > 0 Then
    '    busempact = True
    'End If
    '
    'If InStr(busestposnro, 2) > 0 Then
    '    busempina = True
    'End If
    
    Flog.writeline Espacios(Tabulador * 0) & "Busqueda: " & busnro
    Flog.writeline Espacios(Tabulador * 0) & "Estados de la Busqueda: " & busestposnro
    Flog.writeline Espacios(Tabulador * 0) & "Requerimiento: " & reqpernro
    
    
    Seguir = False
    
    If Not formal Then
            '------------------------------------------------------------------------------------------
            ' Busqueda Informal
            Flog.writeline Espacios(Tabulador * 1) & "Busqueda Informal"
            
            StrSql2 = "SELECT pos_busqueda.buscaracsql, pos_reqbus.reqbusnro "
            StrSql2 = StrSql2 & " FROM pos_busqueda "
            StrSql2 = StrSql2 & " INNER JOIN pos_reqbus ON pos_busqueda.busnro = pos_reqbus.busnro "
            StrSql2 = StrSql2 & " WHERE pos_busqueda.busnro = " & busnro
            'Flog.writeline
            'Flog.writeline Espacios(Tabulador * 1) & "SQL Busqueda: " & StrSql
            OpenRecordset StrSql2, rs
            If Not rs.EOF Then
                caracsql = rs("buscaracsql")
                reqbusnro = rs("reqbusnro")
                Seguir = True
            Else
                Flog.writeline
                Flog.writeline Espacios(Tabulador * 1) & "EOF Busqueda informal"
            End If
            If rs.State = adStateOpen Then rs.Close
    Else
            '------------------------------------------------------------------------------------------
            'Busqueda Formal
            Flog.writeline Espacios(Tabulador * 1) & "Busqueda Formal"
            StrSql2 = "SELECT pos_reqpersonal.reqpercaracsql, pos_reqbus.reqbusnro "
            StrSql2 = StrSql2 & " FROM pos_reqbus "
            StrSql2 = StrSql2 & " INNER JOIN pos_reqpersonal ON pos_reqpersonal.reqpernro = pos_reqbus.reqpernro "
            StrSql2 = StrSql2 & " WHERE pos_reqbus.busnro=" & busnro
            StrSql2 = StrSql2 & " AND pos_reqbus.reqpernro=" & reqpernro
            'Flog.writeline
            'Flog.writeline Espacios(Tabulador * 1) & "SQL Busqueda: " & StrSql
            OpenRecordset StrSql2, rs
            If Not rs.EOF Then
                caracsql = rs("reqpercaracsql")
                reqbusnro = rs("reqbusnro")
                Seguir = True
            Else
                Flog.writeline
                Flog.writeline Espacios(Tabulador * 1) & "EOF Busqueda Formal"
            End If
            If rs.State = adStateOpen Then rs.Close
    End If
    Flog.writeline Espacios(Tabulador * 1) & "SQL Busqueda: " & caracsql
    If Seguir Then
            TiempoAcumulado = GetTickCount
            StrSql = "UPDATE batch_proceso SET bprcprogreso = 5 " & _
                     ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'" & _
                     " WHERE bpronro = " & NroProcesoBatch
            objconnProgreso.Execute StrSql, , adExecuteNoRecords
                     
            Flog.writeline Espacios(Tabulador * 1) & "Elimino los pre-seleccionados si se eligio la opcion eliminarlos"
            '------------------------------------------------------------------------------------------
            ' Elimino los pre-seleccionados si se eligio la opcion eliminarlos
            '------------------------------------------------------------------------------------------
            If Not agregar Then
                StrSql = " DELETE FROM pos_terreqbus "
                StrSql = StrSql & " WHERE reqbusnro = " & reqbusnro
                StrSql = StrSql & " AND conf = 0 "
                objConn.Execute StrSql, , adExecuteNoRecords
            End If
            
            Flog.writeline Espacios(Tabulador * 1) & "Armo la sql conteniendo los terceros"
            '------------------------------------------------------------------------------------------
            ' Armo la sql conteniendo los terceros
            '------------------------------------------------------------------------------------------
            ' Formo la condicion WHERE, excluyendo a los que ya estan seleccionados
            Flog.writeline Espacios(Tabulador * 1) & "Formo la condicion WHERE, excluyendo a los que ya estan seleccionados"
            StrSqlw = StrSqlw & " WHERE " & caracsql & " "
            StrSqlw = StrSqlw & "AND tercero.ternro NOT IN (SELECT ternro "
            StrSqlw = StrSqlw & "FROM pos_terreqbus "
            StrSqlw = StrSqlw & "WHERE reqbusnro = " & reqbusnro & ")"
            
            
            ' NAM - 14/06/2011 - Nuevo Parametro - Controla si debe incluir o no empleados que ya fueron seleccionados en otros busquedas activas.
            If seleccionados = True Then
                StrSqlw = StrSqlw & " AND tercero.ternro NOT IN ("
                StrSqlw = StrSqlw & " SELECT ternro FROM pos_terreqbus pt "
                StrSqlw = StrSqlw & " INNER JOIN pos_reqbus pr ON pt.reqbusnro = pr.reqbusnro "
                StrSqlw = StrSqlw & " INNER JOIN pos_busqueda pb ON pb.busnro = pr.busnro"
                StrSqlw = StrSqlw & " WHERE (pb.busfin <> -1 OR pb.busfin Is Null) AND pt.conf = -1 ) "
            End If
                        
            'reemplazo el token ULTIMO ESTADO por la sql
            StrSqlw = Replace(StrSqlw, "$ULTIMO_ESTADO$", ULTIMO_ESTADO)
            
            'reemplazo el token por el campo de la tabla correspondiente para cada caso
            'si debo buscar la fecha de presentacion del cv del postulante
            StrSqlw_p = Replace(StrSqlw, "$FECHACV$", "pos_postulante.posfecpres")
            'o si debo buscar la fecha de ingreso en caso de empleados
            StrSqlw_e = Replace(StrSqlw, "$FECHACV$", "empleado.empfaltagr")
            
            
            Flog.writeline Espacios(Tabulador * 1) & "Postulantes"
            ' Postulantes
            StrSql = ""
            'If buspos Then
                StrSql = "SELECT tercero.ternro, tercero.terape, tercero.ternom "
                StrSql = StrSql & "FROM tercero "
                StrSql = StrSql & "INNER JOIN pos_postulante ON tercero.ternro = pos_postulante.ternro " 'AND pos_postulante.posest = -1 "
                StrSql = StrSql & "INNER JOIN pos_estpos ON pos_estpos.estposnro = pos_postulante.estposnro AND pos_estpos.estposnro IN (" & busestposnro & ")"
                StrSql = StrSql & StrSqlw_p & "" '<- se usa la _p ya que esta hecho el inner join para la tabla postulantes
            'End If
            
            ' Empleados
            If busempact Or busempina Then
                Flog.writeline Espacios(Tabulador * 1) & "Empleados"
                'If buspos Then
                    StrSql = StrSql & " UNION "
                'End If
                StrSql = StrSql & "SELECT tercero.ternro, tercero.terape, tercero.ternom "
                StrSql = StrSql & "FROM tercero "
                StrSql = StrSql & "INNER JOIN empleado ON tercero.ternro = empleado.ternro "
                If busempact And Not busempina Then
                    StrSql = StrSql & "AND empest = -1 "
                ElseIf Not busempact And busempina Then
                    StrSql = StrSql & "AND empest = 0 "
                End If
                
                'StrSqlw = Replace(StrSqlw, "tercero.ternro IN (SELECT ternro FROM pos_postulante WHERE procnro=", "NOT tercero.ternro IN (SELECT ternro FROM pos_postulante WHERE procnro=")
                'StrSqlw = Replace(StrSqlw, "tercero.ternro IN (SELECT ternro FROM pos_postulante WHERE posrempre", "NOT tercero.ternro IN (SELECT ternro FROM pos_postulante WHERE posrempre")
                StrSqlw_e = Replace(StrSqlw_e, "tercero.ternro IN (SELECT ternro FROM pos_postulante WHERE procnro=", "NOT tercero.ternro IN (SELECT ternro FROM pos_postulante WHERE procnro=")
                StrSqlw_e = Replace(StrSqlw_e, "tercero.ternro IN (SELECT ternro FROM pos_postulante WHERE posrempre", "NOT tercero.ternro IN (SELECT ternro FROM pos_postulante WHERE posrempre")
                
                StrSql = StrSql & StrSqlw_e & "" '<- se usa la _e ya que esta hecho el inner join para la tabla empleado
            End If
            Flog.writeline Espacios(Tabulador * 1) & "SQL generado: " & StrSql
            OpenRecordset StrSql, rs
            
            ' Cuento la cantidad de empleados y postulantes que se agregaran
            If Not rs.EOF Then
                CEmpleadosAProc = rs.RecordCount
                Flog.writeline Espacios(Tabulador * 1) & "Cantidad de Postulantes: " & CEmpleadosAProc
            Else
                Flog.writeline Espacios(Tabulador * 1) & "No hay Postulantes/Empleados que cumplan las condiciones del SQL."
                CEmpleadosAProc = 1
            End If
            ' Defino un número para incrementar
            IncPorc = 95 / CEmpleadosAProc
             
            'Inserto los terceros (empleados/postulantes)
            progresoInc = 0
            Progreso = 5
            Do Until rs.EOF
                If agregar Then
                    StrSql = "SELECT * FROM pos_terreqbus "
                    StrSql = StrSql & " WHERE ternro = " & rs("ternro")
                    StrSql = StrSql & " AND reqbusnro = " & reqbusnro
                    If rs_Pos_Terreqbus.State = adStateOpen Then rs_Pos_Terreqbus.Close
                    OpenRecordset StrSql, rs_Pos_Terreqbus
                    If rs_Pos_Terreqbus.EOF Then
                        StrSql = "INSERT INTO pos_terreqbus "
                        StrSql = StrSql & "(ternro, reqbusnro, conf) "
                        StrSql = StrSql & "VALUES (" & rs("ternro") & "," & reqbusnro & ", 0)"
                        'Flog.writeline Espacios(Tabulador * 1) & "Inserto los terceros (empleados/postulantes)"
                        objConn.Execute StrSql, , adExecuteNoRecords
                        Flog.writeline Espacios(Tabulador * 1) & "Insertando: " & rs("terape") & " " & rs("ternom")
                    Else
                        Flog.writeline Espacios(Tabulador * 1) & "Ya estaba en la busqueda: " & rs("terape") & " " & rs("ternom")
                    End If
                Else
                    StrSql = "INSERT INTO pos_terreqbus "
                    StrSql = StrSql & "(ternro, reqbusnro, conf) "
                    StrSql = StrSql & "VALUES (" & rs("ternro") & "," & reqbusnro & ", 0)"
                    'Flog.writeline Espacios(Tabulador * 1) & "Inserto los terceros (empleados/postulantes)"
                    objConn.Execute StrSql, , adExecuteNoRecords
                    Flog.writeline Espacios(Tabulador * 1) & "Insertando: " & rs("terape") & " " & rs("ternom")
                End If
                rs.MoveNext
                
                
                Flog.writeline Espacios(Tabulador * 1) & "Actualizo el progreso"
                Progreso = Progreso + IncPorc
                progresoInc = progresoInc + IncPorc
                If progresoInc >= 1 Then
                    TiempoAcumulado = GetTickCount
                    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & _
                             ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'" & _
                             " WHERE bpronro = " & NroProcesoBatch
                    objconnProgreso.Execute StrSql, , adExecuteNoRecords
                    progresoInc = 0
                End If
            Loop
            If rs.State = adStateOpen Then rs.Close
        End If
Else
    Flog.writeline Espacios(Tabulador * 1) & "Busqueda de postulantes " & busnro & " no encontrada "
    Flog.writeline Espacios(Tabulador * 1) & "SQL " & StrSql
End If
 
rs_pos_busqueda.Close


MyCommitTrans

If rs_pos_busqueda.State = adStateOpen Then rs_pos_busqueda.Close
If rs.State = adStateOpen Then rs.Close
If rs_Pos_Terreqbus.State = adStateOpen Then rs_Pos_Terreqbus.Close

Set rs_pos_busqueda = Nothing
Set rs = Nothing
Set rs_Pos_Terreqbus = Nothing

End Sub

Public Sub LevantarParametros(ByVal bpronro As Long, ByVal parametros As String)
' --------------------------------------------------------------------------------------------
' Descripcion: Procedimiento para levantar los parametros pasados en batch_proceso en bprcparam
' Autor      : Fernando Favre
' Fecha      : 17/05/2005
' Ult. Mod   :
' Fecha      :
' --------------------------------------------------------------------------------------------
Dim pos1 As Integer
Dim pos2 As Integer
Dim Separador As String

Dim busnro As Long
Dim formal As Boolean
Dim reqpernro As Long
Dim agregar As Boolean
Dim seleccionados As Boolean

Dim params


'Orden de los parametros
'busnro
'formal
'reqpernro
'agregar
'seleccionados

Separador = "@"
' Levanto cada parametro por separado

params = Split(parametros, Separador)
If (UBound(params) <> -1) Then

        busnro = params(0)
        
        formal = params(1)

        reqpernro = params(2)
        
        agregar = params(3)
        
        If UBound(params) = 4 Then
            seleccionados = params(4)
        Else
            seleccionados = False
        End If

End If
Call Generacion(busnro, formal, reqpernro, agregar, seleccionados)
End Sub

