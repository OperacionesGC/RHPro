Attribute VB_Name = "reppresentismo"
Option Explicit

'Const Version = 1.01
'Const FechaVersion = "02/11/2007"
'Modificaciones: Diego Rosso
'Version Inicial

Const Version = 1.02
Const FechaVersion = "19/08/2009"
'Modificaciones: Martin Ferraro - Encriptacion de string connection

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
Global iduser As String
Global horasDia(24) As Boolean
Global tituloReporte As String

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
Dim rscabecera As New ADODB.Recordset
Dim PID As String
Dim arrPronro
Dim TiempoInicialProceso

Dim TiempoAcumulado
Dim totalEmpleados
Dim cantRegistros
Dim Ternro As Long
Dim ListaPar
Dim ArrParametros
Dim pos1
Dim pos2
Dim Parametros
Dim IDCabecera

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
    
    Nombre_Arch = PathFLog & "Rep_Presentismo" & "-" & NroProceso & ".log"
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Nombre_Arch, True)
    
    Flog.writeline "Inicio Proceso Reporte Presentismo : " & Now
    Flog.writeline "Cambio el estado del proceso a Procesando"
    
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
    
    TiempoInicialProceso = GetTickCount
    tituloReporte = ""
    depurar = False
    HuboErrores = False
    
    
    'Cambio el estado del proceso a Procesando
    TiempoAcumulado = GetTickCount
    
    StrSql = "UPDATE batch_proceso SET bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "', bprcpid = " & PID & ", bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcestado = 'Procesando', bprcprogreso = 0 WHERE bpronro = " & NroProceso
    objConn.Execute StrSql, , adExecuteNoRecords
    
    Flog.writeline "Obtengo los datos del proceso"
       
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE bpronro = " & NroProceso
    OpenRecordset StrSql, objRs
    
    If Not objRs.EOF Then
        
        iduser = objRs!iduser
        
        'Cargo los parametros
        Parametros = objRs("bprcparam")
        If Not IsNull(Parametros) Then
            If Len(Parametros) >= 1 Then
                
                'TITULO
                '-----------------------------------------------------------
                pos1 = 1
                pos2 = InStr(pos1, Parametros, "@") - 1
                tituloReporte = Mid(Parametros, pos1, pos2)
                Flog.writeline "Posicion 1 = " & pos1
                Flog.writeline "Pos 2 = " & pos2
                Flog.writeline "Parametro Titulo = " & tituloReporte
                Flog.writeline
                '-------------------------------------------------------------
                
                'IDCabecera
                '------------------------------------------------------------------------------------
                pos1 = pos2 + 2
                pos2 = Len(Parametros)
                IDCabecera = Mid(Parametros, pos1, pos2 - pos1 + 1)
                Flog.writeline "Posicion 1 = " & pos1
                Flog.writeline "Pos 2 = " & pos2
                Flog.writeline "Parametro Numero de Cabecera = " & IDCabecera
                Flog.writeline
                '------------------------------------------------------------------------------------
              
              ' Flog.writeline "Reporte: " & tituloReporte
                           
               
               'EMPIEZA EL PROCESO
               '------------------------------------
                'Obtengo los datos de la cabecera    rscabecera
                 StrSql = "SELECT * FROM  gti_cabpres WHERE gcpnro = " & IDCabecera
                 
                 OpenRecordset StrSql, rscabecera
                 
                 'Obtengo los empleados rsEmpl
                 If rscabecera!Todos = -1 Then
                     StrSql = " SELECT * FROM  empleado WHERE empest = -1 "
                     OpenRecordset StrSql, rsEmpl
                 
                 Else 'Si no traigo los empleados seleccionados en el detalle
                     StrSql = "SELECT ternro FROM  gti_detpres WHERE gcpnro = " & IDCabecera
                     OpenRecordset StrSql, rsEmpl
                 
                 End If
                   
               
               totalEmpleados = rsEmpl.RecordCount
               cantRegistros = totalEmpleados
               
               'Genero por cada empleado
               Do Until rsEmpl.EOF
                  EmpErrores = False
                  Ternro = rsEmpl!Ternro
                  
                  Flog.writeline "Generando datos para el empleado " & Ternro
                  Call generarDatos(Ternro, rscabecera!Fecha, rscabecera!horadesde, rscabecera!horahasta, IDCabecera)
                  
                  cantRegistros = cantRegistros - 1
                  
                  TiempoAcumulado = GetTickCount
                  
                  StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Fix(((totalEmpleados - cantRegistros) * 100) / totalEmpleados) & _
                              ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'" & _
                              ", bprcempleados ='" & CStr(cantRegistros) & "' WHERE bpronro = " & NroProceso
                        
                  objConn.Execute StrSql, , adExecuteNoRecords
                    
                  Flog.writeline "      " & Now & " Fin Procesamiento Empleado "
                  'Paso al siguiente empleado
                  rsEmpl.MoveNext
               Loop
            End If 'bprcparam no es vacio
        Else 'bprcparam es nulo
            Flog.writeline "parametros nulos"
        End If
        Flog.writeline
        Flog.writeline Espacios(Tabulador * 0) & "Terminó de levantar los parametros "
        Flog.writeline

    Else ' Si no hay un proceso pendiente
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
'terno = empleado
'fecha = Fecha de la cabecera de presentismo
'horadesde= hora desde de la cabecera de presentismo
'horahasta= hora hasta de la cabecera de presentismo
Sub generarDatos(ByVal Ternro As Long, ByVal Fecha As Date, ByVal horadesde As String, ByVal horahasta As String, ByVal IDCabecera As Long)

Dim StrSql As String
Dim rsConsult As New ADODB.Recordset
Dim objRsJustif As New ADODB.Recordset
Dim Nombre As String
Dim apellido As String
Dim Legajo As Long
Dim estrnro As Long
Dim CentroCosto As String

Dim objBTurno As New BuscarTurno
Dim objBDia As New BuscarDia
Dim objFeriado As New Feriado
Dim objFechasHoras As New FechasHoras

Dim Hteorico As String
Dim Hreal As String
Dim Licencia  As String 'Descripcion de la justificacion
Dim Estado As Integer ' 1-Presente  2-Ausente  3-Justificado
Dim X As Integer
Dim Cant As Integer
Dim LicDias As Integer  'Cantidad de dias de la justificacion
Dim LicRestantes As Integer 'Dias restanes
Dim LicInicio  'Fecha de inicio de la justificacion
Dim Tomados As Integer  'Cantidad de dias tomados de la justificacion
Dim esFeriado As Boolean 'Si es feriado para el empleado en la fecha a procesar
Dim ExisteReg As Boolean 'Tiene registraciones en el dia dentro del rango horario


On Error GoTo MError

'Inicializo Variables
Licencia = ""
LicDias = 0
LicRestantes = 0
LicInicio = 0
Tomados = 0
LicInicio = ""


'------------------------------------------------------------------
'Busco los datos Basicos del empleado
'------------------------------------------------------------------
StrSql = " SELECT empleg,terape,terape2,ternom,ternom2 "
StrSql = StrSql & " FROM empleado "
StrSql = StrSql & " WHERE ternro= " & Ternro
       
OpenRecordset StrSql, rsConsult

If Not rsConsult.EOF Then
   Nombre = rsConsult!ternom & " " & rsConsult!ternom2
   apellido = rsConsult!terape & " " & rsConsult!terape2
   Legajo = rsConsult!empleg
Else
   Flog.writeline "Error al obtener los datos del empleado"
   GoTo MError
End If

'------------------------------------------------------------------
'Busco el centro de costo del empleado
'------------------------------------------------------------------

StrSql = " SELECT estructura.estrdabr, estructura.estrnro "
StrSql = StrSql & " From his_estructura"
StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro=his_estructura.estrnro AND htetdesde <= " & ConvFecha(Fecha) & " AND (htethasta IS NULL OR htethasta >= " & ConvFecha(Fecha) & ") AND his_estructura.tenro=5 AND his_estructura.ternro=" & Ternro
OpenRecordset StrSql, rsConsult

If Not rsConsult.EOF Then
   CentroCosto = rsConsult!estrdabr
   estrnro = rsConsult!estrnro
Else
   CentroCosto = ""
   estrnro = 0
   Flog.writeline "No se encontraron los datos del centro de costo"
'   GoTo MError
End If


'Comienza el procesamiento del dia
'*********************************

'Busco Feriado
Set objFeriado.Conexion = objConn
esFeriado = objFeriado.Feriado(Fecha, Ternro, True)



'Busco el turno
Set objBTurno.Conexion = objConn
objBTurno.Buscar_Turno Fecha, Ternro, False
    
    
initVariablesTurno objBTurno

    
If Not objBTurno.tiene_turno Then
    Flog.writeline "      " & Now & " sin turno "
Else
    
    'POLITICA 14
    
    
      'Busco el DIA
      Set objBDia.Conexion = objConn
       objBDia.Buscar_Dia Fecha, Fecha, Nro_Turno, Ternro, P_Asignacion, depurar
        'Inicializo variables
       initVariablesDia objBDia
   
       
        StrSql = "SELECT * FROM gti_dias WHERE dianro = " & Nro_Dia
        OpenRecordset StrSql, objRs
    
        If Not objRs.EOF Then
            Hteorico = objRs!diahoradesde1
        Else
            Exit Sub
        End If
End If


'Si no cae dentro de la franja horaria a procesar salgo del procedimiento
If (Hteorico <= horadesde) Or (Hteorico >= horahasta) Or Dia_Libre Then
    Flog.writeline "      " & Now & " El horario teorico no cae dentro de la ventana del reporte "
    Exit Sub
Else
        'Busco registraciones para el dia que se esta procesando
        StrSql = "SELECT regnro,regfecha,reghora, regentsal FROM gti_registracion WHERE " & _
                 "(ternro = " & Ternro & ") AND ( regfecha = " & ConvFecha(Fecha) & ") " & _
                  " and reghora >= '" & horadesde & "' and  reghora <='" & horahasta & "'"
        OpenRecordset StrSql, rsConsult
        
        Cant = 1
        Hreal = ""
        
       If Not rsConsult.EOF Then 'Se encontraron registraciones
             Do While Not rsConsult.EOF
                If Cant Mod 2 <> 0 Then   'Si es impar es una entrada
                    Hreal = rsConsult!reghora
                End If
                Cant = Cant + 1
                rsConsult.MoveNext
             Loop
             
             If Hreal <> "" Then
                Estado = 1 'Presente
                ExisteReg = True
             Else
                 Estado = 2 'Ausente
                 ExisteReg = False
             End If

       End If 'Se encontraron registraciones
       
       If Tiene_Justif = True And Not ExisteReg Then 'Tiene Justificacion y no tiene registraciones
            'Busco la justificacion
            StrSql = "SELECT * FROM gti_justificacion WHERE jusnro= " & nro_justif
            OpenRecordset StrSql, objRsJustif
            
            If objRsJustif!juseltipo = 1 Then  'Dia completo
                Select Case objRsJustif!jussigla
                    Case "NOV"
                        'j_tipo = "NOVEDAD"
                        StrSql = "SELECT gti_novedad.gnovdesabr, gti_novedad.gnovdesde, gti_novedad.gnovhasta FROM gti_novedad WHERE"
                        StrSql = StrSql & " (gnovnro = " & objRsJustif!juscodext & ")"
                        OpenRecordset StrSql, rsConsult
                        If Not objRs.EOF Then
                            Licencia = rsConsult!gnovdesabr
                            LicDias = DateDiff("d", rsConsult!gnovdesde, rsConsult!gnovhasta) + 1
                            If rsConsult!gnovhasta <= Date Then
                                LicRestantes = 0
                            Else
                                LicRestantes = DateDiff("d", rsConsult!gnovhasta, Date) + 1
                            End If
                            If (CDate(rsConsult!gnovdesde) >= Date) Or (CDate(rsConsult!gnovhasta) <= Date) Then
                              Tomados = LicDias
                            Else
                                Tomados = DateDiff("d", rsConsult!gnovdesde, CDate("6/10/2007")) + 1
                            End If
                            If Tomados < 0 Then Tomados = 0
                            LicInicio = rsConsult!gnovdesde
                        End If
                                        
                    Case "LIC"
                        'j_tipo = "LICENCIA"
                        StrSql = "SELECT * FROM emp_lic "
                        StrSql = StrSql & " INNER JOIN tipdia ON emp_lic.tdnro = tipdia.tdnro "
                        StrSql = StrSql & "  WHERE (emp_licnro = " & objRsJustif!juscodext & ")"
                        
                        'StrSql = StrSql & " AND (emp_lic.licestnro = 2)" 'Autorizada
                        OpenRecordset StrSql, rsConsult
                        If Not objRs.EOF Then
                            Licencia = rsConsult!tddesc
                            LicDias = DateDiff("d", rsConsult!elfechadesde, rsConsult!elfechahasta) + 1
                            If rsConsult!elfechahasta <= Date Then
                                LicRestantes = 0
                            Else
                                LicRestantes = DateDiff("d", rsConsult!elfechahasta, Date) + 1
                            End If
                            If (CDate(rsConsult!elfechadesde) >= Date) Or (CDate(rsConsult!elfechahasta) <= Date) Then
                                Tomados = LicDias
                            Else
                                Tomados = DateDiff("d", rsConsult!elfechadesde, CDate("6/10/2007")) + 1
                            End If
                            If Tomados < 0 Then Tomados = 0
                            
                            
                            LicInicio = rsConsult!elfechadesde
                        End If
                    
                    Case "CUR"
                       ' j_tipo = "CURSO"
                        Licencia = "Curso"
                        LicDias = 0
                        LicRestantes = 0
                        LicInicio = Date
                    Case "SUS"
                        'j_tipo = "SUSPENCION"
                        Licencia = "Suspendido por " & CStr(objRsJustif!juscanths) & " hs"
                        LicDias = 0
                        LicRestantes = 0
                        LicInicio = Date
                    End Select
                    Estado = 3 'Justificado
            Else  'Justificacion Parcial - No hay reg
                       
                If Not Dia_Libre And Not esFeriado Then
                    Estado = 2 'Ausente
                Else
                    Exit Sub
                End If
            End If
   Else 'NO TIENE JUSTIFICACION O TIENE REGISTRACIONES
         
       If Not ExisteReg Then
            Estado = 2 'Ausente
        End If
   End If 'Tiene justificacion
      
      '*************************************************************
      'INSERT en gti_rep_presentismo
      '*************************************************************
      StrSql = "INSERT INTO  gti_rep_presentismo"
      StrSql = StrSql & "(bpronro,Fecha,Iduser,gpcnro,ternro,empleg,nombre,apellido, "
      StrSql = StrSql & " entteo,entreal,licencia,estado,licdias,liddiasrest,licinicio,estrnro,estrdabr,titulo, lictomados) "
      StrSql = StrSql & " VALUES "
      StrSql = StrSql & "(" & NroProceso & ", " & ConvFecha(Date) & ", '" & iduser & "', " & IDCabecera
      StrSql = StrSql & ", " & Ternro & ", " & Legajo & ", '" & Nombre & "', '" & apellido & "',"
      StrSql = StrSql & "'" & Hteorico & "', '" & Hreal & "', '" & Licencia & "', " & Estado & "," & LicDias & "," & LicRestantes & ","
      If LicInicio = "" Then
            StrSql = StrSql & "null"
      Else
            StrSql = StrSql & ConvFecha(LicInicio)
      End If
      StrSql = StrSql & "," & estrnro & ",'" & CentroCosto & "','" & tituloReporte & "'," & Tomados & ")"
      'Ejecuto
      objConn.Execute StrSql, , adExecuteNoRecords
      
      Flog.writeline "      Se grabo el registro"
            
End If 'Si no cae dentro de la franja horaria a procesar salgo del procedimiento

Exit Sub

MError:
    Flog.writeline "Error en empleado: " & Legajo & " Error: " & Err.Description
    HuboErrores = True
    EmpErrores = True
    Exit Sub

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


