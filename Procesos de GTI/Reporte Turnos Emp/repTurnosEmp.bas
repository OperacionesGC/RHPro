Attribute VB_Name = "repTurnosEmp"
Option Explicit

'Const Version = 1.01
'Const FechaVersion = "19/10/2006"
''Modificaciones: FGZ
''           Version Inicial

Global Const Version = "1.02" ' Cesar Stankunas
Global Const FechaVersion = "05/08/2009"
Global Const UltimaModificacion = ""    'Encriptacion de string connection

'---------------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------
Dim fs, f

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

Global tiene_turno As Boolean
Global Nro_Turno As Long
Global Tipo_Turno As Integer

Global Tiene_Justif As Boolean
Global nro_justif As Long
Global justif_turno As Boolean
Global p_turcomp As Boolean
Global nro_grupo As Long
Global Nro_fpgo As Integer
Global Fecha_Inicio As Date
Global P_Asignacion  As Boolean
Global Trabaja     As Boolean ' Indica si trabaja para ese dia
Global Orden_Dia As Integer
Global Nro_Dia As Integer
Global Nro_Subturno As Integer
Global Dia_Libre As Boolean
Global Dias_trabajados As Integer
Global Dias_laborables As Integer

Global EmpErrores As Boolean

Global horasDia(24) As Boolean

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
Dim fechadesde
Dim fechahasta
Dim Fecha
Dim tipoDepuracion
Dim historico As Boolean
Dim param
Dim rsEmpl As New ADODB.Recordset
Dim PID As String
Dim arrPronro
Dim TiempoInicialProceso
Dim tituloReporte
Dim TiempoAcumulado
Dim totalEmpleados
Dim cantRegistros
Dim Ternro As Long
Dim ListaPar
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
    
    Call CargarConfiguracionesBasicas
    
    Nombre_Arch = PathFLog & "RepTurnosEmp" & "-" & NroProceso & ".log"
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Nombre_Arch, True)
    
    Flog.writeline "Inicio Proceso Reporte Turnos Empleados : " & Now
    
    'Obtengo el Process ID
    PID = GetCurrentProcessId
    Flog.writeline "-------------------------------------------------"
    Flog.writeline "Version                  : " & Version
    Flog.writeline "Fecha Ultima Modificacion: " & FechaVersion
    Flog.writeline "PID                      : " & PID
    Flog.writeline "-------------------------------------------------"
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 0) & "Inicio :" & Now
    
    On Error Resume Next
    OpenConnection strconexion, objConn
    If Err.Number <> 0 Or Error_Encrypt Then
        Flog.writeline Espacios(Tabulador * 0) & "Problemas en la conexion"
        Exit Sub
    End If
    tituloReporte = ""
    depurar = False
    HuboErrores = False
    
    TiempoInicialProceso = GetTickCount
    
    Flog.writeline "Cambio el estado del proceso a Procesando"
    
    'Cambio el estado del proceso a Procesando
    TiempoAcumulado = GetTickCount
    
    StrSql = "UPDATE batch_proceso SET bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "', bprcpid = " & PID & ", bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcestado = 'Procesando', bprcprogreso = 0 WHERE bpronro = " & NroProceso
    objConn.Execute StrSql, , adExecuteNoRecords
    
    Flog.writeline "Obtengo los datos del proceso"
       
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE bpronro = " & NroProceso
    OpenRecordset StrSql, objRs
    
    If Not objRs.EOF Then
    
       'Obtengo los parametros del proceso
       tituloReporte = objRs!bprcparam
       fechadesde = CDate(objRs!bprcfecdesde)
       fechahasta = CDate(objRs!bprcfechasta)
       
       Flog.writeline "Reporte: " & tituloReporte
       Flog.writeline "En el rango de Fechas: " & fechadesde & "-" & fechahasta
       
       'EMPIEZA EL PROCESO
       
       'Obtengo los empleados
       CargarEmpleados NroProceso, rsEmpl
       
       totalEmpleados = rsEmpl.RecordCount
       cantRegistros = totalEmpleados
       
       'Genero por cada empleado/fecha los horarios
       Do Until rsEmpl.EOF
          EmpErrores = False
          Ternro = rsEmpl!Ternro
          
          Flog.writeline "Generando datos para el empleado " & Ternro
          Call generarDatosTurnoFecha(Ternro, tituloReporte, fechadesde, fechahasta)
          
          cantRegistros = cantRegistros - 1
          
          TiempoAcumulado = GetTickCount
          
          StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Fix(((totalEmpleados - cantRegistros) * 100) / totalEmpleados) & _
                      ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'" & _
                      ", bprcempleados ='" & CStr(cantRegistros) & "' WHERE bpronro = " & NroProceso
                
          objConn.Execute StrSql, , adExecuteNoRecords
          
          'Si se generaron todos los recibos de sueldo del empleado correctamente lo borro
          If Not EmpErrores Then
              StrSql = " DELETE FROM batch_empleado "
              StrSql = StrSql & " WHERE bpronro = " & NroProceso
              StrSql = StrSql & " AND ternro = " & Ternro
    
              objConn.Execute StrSql, , adExecuteNoRecords
          End If
          
          rsEmpl.MoveNext
       Loop
    
    Else
        Exit Sub
    End If
   
    'Actualizo el estado del proceso
    If Not HuboErrores Then
       StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado' WHERE bpronro = " & NroProceso
    Else
       StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Incompleto' WHERE bpronro = " & NroProceso
    End If
    
    objConn.Execute StrSql, , adExecuteNoRecords
    
    Flog.writeline "Fin :" & Now
    Flog.Close

    Exit Sub
    
CE:
    HuboErrores = True
    Flog.writeline " Error: " & Err.Description & Now

End Sub

'--------------------------------------------------------------------
' Se encarga de generar los datos
'--------------------------------------------------------------------
Sub generarDatosTurnoFecha(Ternro As Long, tituloReporte, fechadesde, fechahasta)

Dim StrSql As String
Dim rsConsult As New ADODB.Recordset
Dim Nombre
Dim apellido
Dim Legajo
Dim empfecalta
Dim Fecha As Date

On Error GoTo MError

'------------------------------------------------------------------
'Busco los datos del empleado
'------------------------------------------------------------------
StrSql = " SELECT empleg,terape,terape2,ternom,ternom2,empfecalta,empremu "
StrSql = StrSql & " FROM empleado "
StrSql = StrSql & " WHERE ternro= " & Ternro
       
OpenRecordset StrSql, rsConsult

If Not rsConsult.EOF Then
   Nombre = rsConsult!ternom & " " & rsConsult!ternom2
   apellido = rsConsult!terape & " " & rsConsult!terape2
   Legajo = rsConsult!empleg
   empfecalta = rsConsult!empfecalta
Else
   Flog.writeline "Error al obtener los datos del empleado"
   GoTo MError
End If

'------------------------------------------------------------------
'Armo la SQL para guardar los datos
'------------------------------------------------------------------

StrSql = " INSERT INTO gti_rep_turnos_emp "
StrSql = StrSql & "( bprcnro , Ternro, fechadesde, fechahasta,"
StrSql = StrSql & "  terape , ternom, descripcion) "
StrSql = StrSql & " VALUES"
StrSql = StrSql & "(" & NroProceso
StrSql = StrSql & "," & Ternro
StrSql = StrSql & "," & ConvFecha(fechadesde)
StrSql = StrSql & "," & ConvFecha(fechahasta)
StrSql = StrSql & ",'" & apellido & "'"
StrSql = StrSql & ",'" & Nombre & "'"
StrSql = StrSql & ",'" & tituloReporte & "'"
StrSql = StrSql & ")"

'------------------------------------------------------------------
'Guardo los datos en la BD
'------------------------------------------------------------------
objConn.Execute StrSql, , adExecuteNoRecords

'Para cada fecha genero una entrada en la BD
Fecha = CDate(fechadesde)
          
Do Until DateDiff("d", Fecha, CDate(fechahasta)) < 0

   Call generarDatosFecha(Ternro, Fecha)
   
   Fecha = DateAdd("d", 1, Fecha)
Loop
    
Exit Sub

MError:
    Flog.writeline "Error en empleado: " & Legajo & " Error: " & Err.Description
    HuboErrores = True
    EmpErrores = True
    Exit Sub

End Sub


'--------------------------------------------------------------------
' Se encarga de generar un ResultSet de los empleados a cambiar
' si el RS es vacio significa que hay que aplicarlo sobre todos
'--------------------------------------------------------------------
Sub CargarEmpleados(NroProc, ByRef rsEmpl As ADODB.Recordset)

Dim StrEmpl As String

    StrEmpl = "SELECT * FROM batch_empleado WHERE bpronro = " & NroProc
    
    OpenRecordset StrEmpl, rsEmpl
End Sub

Private Sub initVariablesTurno(ByRef T As BuscarTurno)
   p_turcomp = T.Compensa_Turno
   nro_grupo = T.Empleado_Grupo
   nro_justif = T.Justif_Numero
   justif_turno = T.justif_turno
   Tiene_Justif = T.Tiene_Justif
   Fecha_Inicio = T.FechaInicio
   Nro_fpgo = T.Numero_FPago
   Nro_Turno = T.Turno_Numero
   tiene_turno = T.tiene_turno
   Tipo_Turno = T.Turno_Tipo
   P_Asignacion = T.Tiene_PAsignacion
End Sub

Private Sub initVariablesDia(ByRef D As BuscarDia)
   Dia_Libre = D.Dia_Libre
   Nro_Dia = D.Numero_Dia
   Nro_Subturno = D.SubTurno_Numero
   Orden_Dia = D.Orden_Dia
   Trabaja = D.Trabaja
End Sub

Private Sub setearHorarioTeorico()

Dim desde
Dim hasta
Dim i
    
    StrSql = "SELECT * FROM gti_dias WHERE dianro = " & Nro_Dia
    OpenRecordset StrSql, objRs
    
    If Not objRs.EOF Then
    
        If (objRs!diahoradesde1 <> "0000" Or objRs!diahorahasta1 <> "0000") Then
            desde = CInt(Mid(objRs!diahoradesde1, 1, 1) & Mid(objRs!diahoradesde1, 2, 1))
            hasta = CInt(Mid(objRs!diahorahasta1, 1, 1) & Mid(objRs!diahorahasta1, 2, 1))
            
            For i = desde To hasta
               horasDia(i) = True
            Next
        End If
        
        If (objRs!diahoradesde2 <> "0000" Or objRs!diahorahasta2 <> "0000") Then
            desde = CInt(Mid(objRs!diahoradesde2, 1, 1) & Mid(objRs!diahoradesde2, 2, 1))
            hasta = CInt(Mid(objRs!diahorahasta2, 1, 1) & Mid(objRs!diahorahasta2, 2, 1))
            
            For i = desde To hasta
               horasDia(i) = True
            Next
        End If
        
        If (objRs!diahoradesde3 <> "0000" Or objRs!diahorahasta3 <> "0000") Then
            desde = CInt(Mid(objRs!diahoradesde3, 1, 1) & Mid(objRs!diahoradesde3, 2, 1))
            hasta = CInt(Mid(objRs!diahorahasta3, 1, 1) & Mid(objRs!diahorahasta3, 2, 1))
            
            For i = desde To hasta
               horasDia(i) = True
            Next
        End If
        
    End If
    
    objRs.Close
End Sub



'--------------------------------------------------------------------
' Se encarga de generar los datos
'--------------------------------------------------------------------
Sub generarDatosFecha(Ternro As Long, FechaActual As Date)

Dim StrSql As String
Dim rsConsult As New ADODB.Recordset
Dim Nombre
Dim apellido
Dim Legajo
Dim empfecalta

Dim objBTurno As New BuscarTurno
Dim objBDia As New BuscarDia
Dim objFeriado As New Feriado
Dim objFechasHoras As New FechasHoras
Dim i
Dim Fecha As Date
Dim Laborable

On Error GoTo MError

Fecha = FechaActual

'------------------------------------------------------------------
'Inicializo el arreglo de horas
'------------------------------------------------------------------
For i = 0 To 23
   horasDia(i) = False
Next

'------------------------------------------------------------------
'Busco la informacion del horario teorico
'------------------------------------------------------------------

Set objBTurno.Conexion = objConn
'Set objBTurno.ConexionTraza = CnTraza
    
objBTurno.Buscar_Turno Fecha, Ternro, False
    
initVariablesTurno objBTurno
'Flog.writeline Now & " Bturno"
    
If Not objBTurno.tiene_turno Then
    Flog.writeline Now & " sin turno "
Else
    If objBTurno.Tiene_Justif Then
      Set objBDia.Conexion = objConn
  '   Set objBDia.ConexionTraza = CnTraza
   
       objBDia.Buscar_Dia Fecha, Fecha, Nro_Turno, Ternro, P_Asignacion, depurar
   
       initVariablesDia objBDia
   
       Call setearHorarioTeorico
    Else
       Set objBDia.Conexion = objConn
    '   Set objBDia.ConexionTraza = CnTraza
       
       objBDia.Buscar_Dia Fecha, Fecha_Inicio, Nro_Turno, Ternro, P_Asignacion, depurar
       
       initVariablesDia objBDia
       
       Call setearHorarioTeorico
    End If
End If

If Dia_Libre Then
   Laborable = 0
Else
   Laborable = -1
End If
    
'------------------------------------------------------------------
'Armo la SQL para guardar los datos
'------------------------------------------------------------------

StrSql = " INSERT INTO gti_rep_turemp_det "
StrSql = StrSql & "( bprcnro , Ternro, Fecha, "
StrSql = StrSql & "  turnro,subturnro,laborable, "
StrSql = StrSql & "  hora0 , hora1, hora2,"
StrSql = StrSql & "  hora3 , hora4, hora5,"
StrSql = StrSql & "  hora6 , hora7, hora8,"
StrSql = StrSql & "  hora9 , hora10, hora11,"
StrSql = StrSql & "  hora12 , hora13, hora14,"
StrSql = StrSql & "  hora15 , hora16, hora17,"
StrSql = StrSql & "  hora18 , hora19, hora20,"
StrSql = StrSql & "  hora21 , hora22, hora23)"
StrSql = StrSql & " VALUES"
StrSql = StrSql & "(" & NroProceso
StrSql = StrSql & "," & Ternro
StrSql = StrSql & "," & ConvFecha(Fecha)
StrSql = StrSql & "," & Nro_Turno
StrSql = StrSql & "," & Nro_Subturno
StrSql = StrSql & "," & Laborable

For i = 0 To 23
   If horasDia(i) Then
     StrSql = StrSql & ",-1"
   Else
     StrSql = StrSql & ",0"
   End If
Next

StrSql = StrSql & ")"

'------------------------------------------------------------------
'Guardo los datos en la BD
'------------------------------------------------------------------
objConn.Execute StrSql, , adExecuteNoRecords

Exit Sub

MError:
    Flog.writeline "Error en empleado: " & Legajo & " Error: " & Err.Description
    HuboErrores = True
    EmpErrores = True
    Exit Sub

End Sub

