Attribute VB_Name = "repAuditoria"
Option Explicit

'Version: 1.00
'

'Const Version = 1.01
'Const FechaVersion = "26/01/2006" ' Se agregaron string de errores

Global Const Version = "1.01" ' Cesar Stankunas
Global Const FechaVersion = "05/08/2009"
Global Const UltimaModificacion = ""    'Encriptacion de string connection

Dim fs, f
'Global Flog

Global Const Tabulador = 8

Dim NroLinea As Long
Dim crpNro As Long
Dim RegLeidos As Long
Dim RegError As Long
Dim RegFecha As Date
Dim NroProceso As Long

Global Path As String
Global NArchivo As String
Global Rta
Global HuboErrores As Boolean
Global EmpErrores As Boolean

Global taccion As Integer
Global accion As Integer
Global tusuario As Integer
Global usuario As String
Global confnro As String
Global campos As String
Global tipoEmpleados As String
Global fechadesde As Date
Global fechahasta As Date

'DATOS DE LA TABLA batch_proceso
Global bpfecha As Date
Global bphora As String
Global bpusuario As String

Global repNro As Integer
Global conceptos As String
Global acumuladores As String
Global procesos As String
Global idUser As String

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
Dim objRs2 As New ADODB.Recordset
Dim tipoDepuracion
Dim historico As Boolean
Dim param
Dim ternro
Dim rsEmpl As New ADODB.Recordset
Dim I
Dim totalEmpleados
Dim cantRegistros
Dim PID As String
Dim ArrParametros
Dim parametros As String

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
    
    Nombre_Arch = PathFLog & "ReporteAuditoria" & "-" & NroProceso & ".log"
    
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
    OpenConnection strconexion, objConn
    If Err.Number <> 0 Or Error_Encrypt Then
        Flog.writeline Espacios(Tabulador * 0) & "Problemas en la conexion"
        Exit Sub
    End If
    HuboErrores = False
    
    Flog.writeline "Inicio Proceso de Auditoria : " & Now
    Flog.writeline "Cambio el estado del proceso a Procesando"
    
    'Cambio el estado del proceso a Procesando
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcestado = 'Procesando', bprcpid = " & PID & " WHERE bpronro = " & NroProceso
    objConn.Execute StrSql, , adExecuteNoRecords
    
    Flog.writeline "Obtengo los datos del proceso"
    
    TiempoAcumulado = GetTickCount
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE bpronro = " & NroProceso
    OpenRecordset StrSql, objRs
    
    If Not objRs.EOF Then
       
       'Obtengo los parametros del proceso
       parametros = objRs!bprcparam
       ArrParametros = Split(parametros, "@")
       
       'Obtengo si se eligieron todas las acciones
       taccion = ArrParametros(0)
       
       'Obtengo la acción elegida
       accion = ArrParametros(1)
       
       'Obtengo si se eligieron todos los usuarios
       tusuario = ArrParametros(2)
       
       'Obtengo el usuario
       usuario = ArrParametros(3)
       
       'Obtengo la configuraciones
       confnro = ArrParametros(4)
       
       'Obtengo los campos que se quieren filtrar
       campos = ArrParametros(5)
       
       'Obtengo los empleados a los cuales aplicar la auditoria
       '0 = todos, sino una lista separada por comas
       tipoEmpleados = ArrParametros(6)
       
       'Obtengo las fechas
       fechadesde = objRs!bprcfecdesde
       fechahasta = objRs!bprcfechasta
       
       
       'Obtengo la fecha del proceso
       bpfecha = objRs!bprcfecha
       
       'Obtengo la hora del proceso
       bphora = objRs!bprchora
       
       'Obtengo el usuario del proceso
       bpusuario = objRs!idUser
                     
       'Obtengo el titulo del reporte
       'tituloReporte = arrParametros(5)
       
    
       'EMPIEZA EL PROCESO

       Call generarAuditoria
                
 
    Else
        Exit Sub
    End If
   
    'Actualizo el estado del proceso
    If Not HuboErrores Then
       StrSql = "UPDATE batch_proceso SET  bprcprogreso =100, bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado' WHERE bpronro = " & NroProceso
    Else
       StrSql = "UPDATE batch_proceso SET  bprcprogreso =100, bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Incompleto' WHERE bpronro = " & NroProceso
    End If
    
    objConn.Execute StrSql, , adExecuteNoRecords
    
    Flog.writeline "Fin :" & Now
    Flog.Close

    Exit Sub
    
CE:
    HuboErrores = True
    Flog.writeline " Error: " & Err.Description & Now
    Flog.writeline "Ultima SQL Ejecutada: " & StrSql

End Sub

Function controlNull(Str)
  If Trim(Str) = "" Then
     controlNull = "null"
  Else
     controlNull = "'" & Str & "'"
  End If
End Function

'--------------------------------------------------------------------
' Se encarga de buscar las auditorias
'--------------------------------------------------------------------
Sub generarAuditoria()

Dim StrSql As String
Dim rsConsult1 As New ADODB.Recordset
Dim rsConsult2 As New ADODB.Recordset
Dim Legajo As Long
Dim Cantidad As Long
Dim cantidadProcesada As Long
Dim Empleado As String
Dim EmpTernro

On Error GoTo MError

StrSql = " SELECT acdesc, aud_actual, aud_ant, aud_campo, aud_fec, aud_hor, "
StrSql = StrSql & " iduser, aud_ternro, aud_des, aud_rel, aud_campnro "
StrSql = StrSql & " FROM auditoria "
StrSql = StrSql & " INNER JOIN accion ON accion.acnro = auditoria.acnro "
StrSql = StrSql & " WHERE (auditoria.aud_fec >= " & ConvFecha(fechadesde) & " ) AND "
StrSql = StrSql & " (auditoria.aud_fec <= " & ConvFecha(fechahasta) & " ) "

If confnro <> "0" And confnro <> "0,0" Then
   StrSql = StrSql & " AND auditoria.caudnro IN (" & confnro & " ) "
End If
If taccion = 0 Then
   StrSql = StrSql & " AND auditoria.acnro = " & accion
End If
If tusuario = 0 Then
   StrSql = StrSql & " AND auditoria.iduser = '" & usuario & "'"
End If
StrSql = StrSql & " AND auditoria.aud_campnro IN (" & campos & " ) "
If tipoEmpleados <> "0" Then
    StrSql = StrSql & " AND ( ( auditoria.aud_ternro IN (SELECT ternro FROM batch_empleado WHERE batch_empleado.bpronro= " & NroProceso & ") ) "
    StrSql = StrSql & " OR    ( auditoria.aud_rel    IN (SELECT ternro FROM batch_empleado WHERE batch_empleado.bpronro= " & NroProceso & ") ) "
    StrSql = StrSql & "     ) "
End If

StrSql = StrSql & " ORDER BY aud_fec "

OpenRecordset StrSql, rsConsult1

Cantidad = rsConsult1.RecordCount
cantidadProcesada = Cantidad

Progreso = 0
If Cantidad = 0 Then
    Cantidad = 1
End If
IncPorc = 95 / Cantidad

If rsConsult1.EOF Then
    Flog.writeline Espacios(Tabulador * 0) & "No se encontraron registros de Auditoria: " & StrSql
End If

Do Until rsConsult1.EOF

    Empleado = ""
    Legajo = 0
    EmpTernro = 0

    If rsConsult1!aud_ternro <> 0 Then
          
        StrSql = " SELECT empleg,terape,ternom "
        StrSql = StrSql & " FROM empleado "
        StrSql = StrSql & " WHERE empleado.ternro = " & rsConsult1!aud_ternro
             
        OpenRecordset StrSql, rsConsult2
        
        If Not rsConsult2.EOF Then
           EmpTernro = rsConsult1!aud_ternro
           Legajo = rsConsult2!empleg
           Empleado = rsConsult2!empleg & " - " & rsConsult2!terape & ", " & rsConsult2!ternom
        Else
             
                If IsNumeric(rsConsult1!aud_rel) Then
                   StrSql = " SELECT empleg,terape,ternom "
                   StrSql = StrSql & " FROM empleado "
                   StrSql = StrSql & " WHERE empleado.ternro = " & rsConsult1!aud_rel
                
                   OpenRecordset StrSql, rsConsult2
                   
                   If Not rsConsult2.EOF Then
                      EmpTernro = rsConsult1!aud_rel
                      Legajo = rsConsult2!empleg
                      Empleado = rsConsult2!empleg & " - " & rsConsult2!terape & ", " & rsConsult2!ternom
                    End If
                End If
        End If
    Else
           If IsNumeric(rsConsult1!aud_rel) Then
               StrSql = " SELECT empleg,terape,ternom "
               StrSql = StrSql & " FROM empleado "
               StrSql = StrSql & " WHERE empleado.ternro = " & rsConsult1!aud_rel
                
               OpenRecordset StrSql, rsConsult2
               
               If Not rsConsult2.EOF Then
                  EmpTernro = rsConsult1!aud_rel
                  Legajo = rsConsult2!empleg
                  Empleado = rsConsult2!empleg & " - " & rsConsult2!terape & ", " & rsConsult2!ternom
                End If
            End If
    End If
     
    StrSql = " INSERT INTO rep_auditoria "
    StrSql = StrSql & " (bpronro , bpro_fecha, bpro_hora, bpro_iduser, aud_fec, aud_hor, "
    StrSql = StrSql & " aud_iduser, aud_des, aud_actual, aud_ant, ternro, aud_campo, "
    StrSql = StrSql & " acc_desc, empleado, fecha_desde, fecha_hasta, aud_campnro )"
    StrSql = StrSql & " VALUES "
    StrSql = StrSql & "(" & NroProceso
    StrSql = StrSql & "," & ConvFecha(bpfecha)
    StrSql = StrSql & ",'" & bphora & "'"
    StrSql = StrSql & ",'" & bpusuario & "'"
    StrSql = StrSql & "," & ConvFecha(rsConsult1!aud_fec)
    StrSql = StrSql & ",'" & rsConsult1!aud_hor & "'"
    StrSql = StrSql & ",'" & rsConsult1!idUser & "'"
    StrSql = StrSql & ",'" & rsConsult1!aud_des & "'"
    StrSql = StrSql & ",'" & rsConsult1!aud_actual & "'"
    StrSql = StrSql & ",'" & rsConsult1!aud_ant & "'"
    StrSql = StrSql & "," & EmpTernro
    StrSql = StrSql & ",'" & rsConsult1!aud_campo & "'"
    StrSql = StrSql & ",'" & rsConsult1!acdesc & "'"
    StrSql = StrSql & ",'" & Empleado & "'"
    StrSql = StrSql & "," & ConvFecha(fechadesde)
    StrSql = StrSql & "," & ConvFecha(fechahasta)
    StrSql = StrSql & "," & rsConsult1!aud_campnro
    StrSql = StrSql & ")"
    
    objConn.Execute StrSql, , adExecuteNoRecords
    
    'Actualizo el estado del proceso
    TiempoAcumulado = GetTickCount
    Progreso = Progreso + IncPorc
          
    cantidadProcesada = cantidadProcesada - 1
          
    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & _
             ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'" & _
             ", bprcempleados ='" & CStr(cantidadProcesada) & "' WHERE bpronro = " & NroProceso
             
    objConn.Execute StrSql, , adExecuteNoRecords

    rsConsult1.MoveNext
    
Loop

rsConsult1.Close

Exit Sub

MError:
    Flog.writeline "Error: " & Err.Description
    Flog.writeline "Ultima SQL Ejecutada: " & StrSql
    HuboErrores = True
    EmpErrores = True
    Exit Sub
End Sub


Function numberForSQL(Str)
   
  numberForSQL = Replace(Str, ",", ".")

End Function


Function strForSQL(Str)
   
  If IsNull(Str) Then
     strForSQL = "NULL"
  Else
     strForSQL = Str
  End If

End Function

Function sinDatos(Str)

  If IsNull(Str) Then
     sinDatos = True
  Else
     If Trim(Str) = "" Then
        sinDatos = True
     Else
        sinDatos = False
     End If
  End If

End Function


