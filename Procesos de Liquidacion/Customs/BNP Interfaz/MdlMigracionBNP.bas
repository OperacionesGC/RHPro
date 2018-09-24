Attribute VB_Name = "MdlMigracionBNP"
Option Explicit

'Global Const Version = "1.01"
'Global Const FechaModificacion = "03/03/2006"
'Global Const UltimaModificacion = " "

'Global Const Version = "1.02"
'Global Const FechaModificacion = "23/10/2006"
'Global Const UltimaModificacion = " " 'FAF - Se cambiaron depeve33, depeve34, evedet05 y evedet04. La funcion instr
                                       ' evedet09 no funcionaba para la baja de Expatriados
Global Const Version = "1.03"
Global Const FechaModificacion = "25/10/2006"
Global Const UltimaModificacion = " " 'FAF - Se cambiaron evedet08 y evedet10. Mismo caso que evedet09, no se asignaba la caunro de baja
'-----------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------

Global IdUser As String
Global Fecha As Date
Global hora As String

Global intnro As Integer

Public Sub Main()
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento inicial del Proceso.
' Autor      : Fernando Favre
' Fecha      : 03/03/2006
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
    
    ' carga las configuraciones basicas, formato de fecha, string de conexion,
    ' tipo de BD y ubicacion del archivo de log
    Call CargarConfiguracionesBasicas

    'Abro la conexion
    OpenConnection strconexion, objConn
    OpenConnection strconexion, objconnProgreso
    
    Nombre_Arch = PathFLog & "InterfazBNP" & "-" & NroProcesoBatch & ".log"
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Nombre_Arch, True)
    
    ' Obtengo el Process ID
    PID = GetCurrentProcessId
    Flog.writeline "-----------------------------------------------------------------"
    Flog.writeline "Version = " & Version
    Flog.writeline "Modificacion = " & UltimaModificacion
    Flog.writeline "Fecha = " & FechaModificacion
    Flog.writeline "-----------------------------------------------------------------"
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 0) & "PID = " & PID
    
    'Cambio el estado del proceso a Procesando y el PID
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcprogreso = 0, bprcestado = 'Procesando', bprcpid = " & PID & " WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE btprcnro = 125 AND bpronro =" & NroProcesoBatch
    OpenRecordset StrSql, rs_batch_proceso
    
    TiempoInicialProceso = GetTickCount
    
    If Not rs_batch_proceso.EOF Then
        IdUser = rs_batch_proceso!IdUser
        Fecha = rs_batch_proceso!bprcfecha
        hora = rs_batch_proceso!bprchora
        bprcparam = rs_batch_proceso!bprcparam
        
        rs_batch_proceso.Close
        Set rs_batch_proceso = Nothing
        
        Call LevantarParamteros(NroProcesoBatch, bprcparam)
    End If
    
    TiempoFinalProceso = GetTickCount
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 0) & "=================================================="
    Flog.writeline Espacios(Tabulador * 0) & "Tiempo del proceso (milisegundos): " & (TiempoFinalProceso - TiempoInicialProceso)
    Flog.writeline Espacios(Tabulador * 0) & "=================================================="
    
    If Not HuboError Then
        StrSql = "DELETE FROM batch_empleado WHERE bpronro = " & NroProcesoBatch
        objConn.Execute StrSql, , adExecuteNoRecords
        
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado' WHERE bpronro = " & NroProcesoBatch
    Else
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Error' WHERE bpronro = " & NroProcesoBatch
    End If
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    Flog.Close
    objconnProgreso.Close
    objConn.Close
    
End Sub

Public Sub Generacion(ByVal intnro As Integer)
' --------------------------------------------------------------------------------------------
' Descripcion: Procedimiento de generacion de datos
' Autor      : Fernando Favre
' Fecha      : 07/03/2006
' Ult. Mod   :
' Fecha      :
' --------------------------------------------------------------------------------------------
Dim inttipo As Integer
Dim excluido As Boolean
Dim tiene_evento As Boolean
Dim programa As Boolean
Dim valorfijo As Boolean
Dim v_valor As String
Dim ok_valor As Boolean
Dim v_tipovalor As Integer
Dim v_prog As String
Dim msg_valor As String
Dim nuevo As Boolean

'Registros
Dim rs_interfaz As New ADODB.Recordset
Dim rs_Empleados As New ADODB.Recordset
Dim rs_event As New ADODB.Recordset
Dim rs_Event_Topic As New ADODB.Recordset
Dim rs_topic_field As New ADODB.Recordset
Dim rs_Event_Topic_Field As New ADODB.Recordset
Dim rs_field_value As New ADODB.Recordset

On Error GoTo CE

MyBeginTrans


' Busco los valores de la interfaz
StrSql = "SELECT * FROM interfaz WHERE intnro = " & intnro
OpenRecordset StrSql, rs_interfaz
If rs_interfaz.EOF Then
    Flog.writeline Espacios(Tabulador * 1) & "No se encontraron los valores de la interfaz " & intnro
    GoTo Fin
Else
    inttipo = rs_interfaz!inttipo
End If
rs_interfaz.Close

' Para todos los empleados seleccionados
StrSql = "SELECT * FROM batch_empleado "
StrSql = StrSql & " INNER JOIN empleado ON batch_empleado.ternro = empleado.ternro "
StrSql = StrSql & " WHERE batch_empleado.bpronro = " & NroProcesoBatch
If intnro = 1 Then
    StrSql = StrSql & " AND empleado.empest = -1"
End If
OpenRecordset StrSql, rs_Empleados


' Seteo de las variables de progreso
Progreso = 0
CConceptosAProc = rs_Empleados.RecordCount
If CConceptosAProc = 0 Then
    CConceptosAProc = 1
    Flog.writeline Espacios(Tabulador * 1) & " No hay Empleados para procesar."
End If
IncPorc = (99 / CConceptosAProc)


' Para cada uno de los empleados
Do Until rs_Empleados.EOF
    
    Flog.writeline Espacios(Tabulador * 1) & "Empleado: " & rs_Empleados!empleg & " - " & rs_Empleados!terape & " " & rs_Empleados!terape2 & ", " & rs_Empleados!ternom & " " & rs_Empleados!ternom2
    
    ' Se borran los datos anteriores
    StrSql = "DELETE FROM field_value "
    StrSql = StrSql & "WHERE intnro = " & intnro & " AND ternro = " & rs_Empleados!ternro
    objConn.Execute StrSql, , adExecuteNoRecords

    ' Considerar al empleado como no generado
    StrSql = "DELETE FROM empgenerado "
    StrSql = StrSql & "WHERE intnro = " & intnro & " AND ternro = " & rs_Empleados!ternro
    objConn.Execute StrSql, , adExecuteNoRecords
    
    ' Se borra el LOG
    StrSql = "DELETE FROM intemplog "
    StrSql = StrSql & "WHERE intnro = " & intnro & " AND ternro = " & rs_Empleados!ternro
    objConn.Execute StrSql, , adExecuteNoRecords
    
    ' Busco los eventos asociados a la interface
    StrSql = "SELECT * FROM event "
    StrSql = StrSql & " WHERE eventactive = -1 "
    If inttipo = 1 Then
        StrSql = StrSql & " AND eventtotal = -1 "
    Else
        StrSql = StrSql & " AND eventtotal = 0 "
    End If
    StrSql = StrSql & " ORDER BY eventorder"
    OpenRecordset StrSql, rs_event
    
    ' Para cada uno de los eventos
    Do Until rs_event.EOF
        
        ' Verificar si el evento esta excluido por otro
        Call event_excluido(intnro, 1, CLng(rs_Empleados!ternro), rs_event!eventnro, excluido)
        If excluido Then
            Flog.writeline Espacios(Tabulador * 2) & "EXCLUIDO el Evento: " & rs_event!eventcode & " - " & rs_event!eventdesc & " - por otro."
            GoTo bloque_evento
        End If
        
        Flog.writeline Espacios(Tabulador * 2) & "Evento (programa): " & rs_event!eventcode & " - " & rs_event!eventdesc & " - (" & rs_event!eventverifprg & ")"
        
        ' Verificar si hay programa de determinación
        '   Si existe --> Verificar si el empleado tiene un evento de este tipo
        Select Case rs_event!eventverifprg
            Case "evedet04.p":
                Call evedet04(intnro, 1, CLng(rs_Empleados!ternro), rs_event!eventnro, tiene_evento)
            Case "evedet05.p":
                Call evedet05(intnro, 1, CLng(rs_Empleados!ternro), rs_event!eventnro, tiene_evento)
            Case "evedet06.p":
                Call evedet06(intnro, 1, CLng(rs_Empleados!ternro), rs_event!eventnro, tiene_evento)
            Case "evedet07.p":
                Call evedet07(intnro, 1, CLng(rs_Empleados!ternro), rs_event!eventnro, tiene_evento)
            Case "evedet08.p":
                Call evedet08(intnro, 1, CLng(rs_Empleados!ternro), rs_event!eventnro, tiene_evento)
            Case "evedet09.p":
                Call evedet09(intnro, 1, CLng(rs_Empleados!ternro), rs_event!eventnro, tiene_evento)
            Case "evedet10.p":
                Call evedet10(intnro, 1, CLng(rs_Empleados!ternro), rs_event!eventnro, tiene_evento)
            Case "evedet11.p":
                Call evedet11(intnro, 1, CLng(rs_Empleados!ternro), rs_event!eventnro, tiene_evento)
            Case "evedet12.p":
                Call evedet12(intnro, 1, CLng(rs_Empleados!ternro), rs_event!eventnro, tiene_evento)
            Case "evedet13.p":
                Call evedet13(intnro, 1, CLng(rs_Empleados!ternro), rs_event!eventnro, tiene_evento)
            Case "evedet14.p":
                Call evedet14(intnro, 1, CLng(rs_Empleados!ternro), rs_event!eventnro, tiene_evento)
            Case Else
                Flog.writeline Espacios(Tabulador * 3) & "ERROR. Programa de determinación NO definido para el evento (event.eventverifprg)."
                Call registro_log(intnro, 1, CLng(rs_Empleados!ternro), 2, rs_event!eventcode & ": Programa Detección no encontrado: " & rs_event!eventverifprg)
        End Select
        
        
        If Not tiene_evento Then
            Flog.writeline Espacios(Tabulador * 3) & "No tiene evento."
            GoTo bloque_evento
        End If
        
        ' Recorrer aquellos tópicos que son obligatorios (asoctype = 1)
        ' o aquellos que son opcionales y se ha elegido informarlo (asoctype = 2)
        StrSql = "SELECT event_topic.topicnro, event_topic.eventnro, topic.topicheader, topic.topicname FROM event_topic "
        StrSql = StrSql & " INNER JOIN topic ON event_topic.topicnro = topic.topicnro "
        StrSql = StrSql & " WHERE event_topic.eventnro = " & rs_event!eventnro
        StrSql = StrSql & " AND (event_topic.asoctype = 1 OR event_topic.asoctype = 2)"
        OpenRecordset StrSql, rs_Event_Topic
        
        ' Recorrer los campos de datos para el topico
        Do Until rs_Event_Topic.EOF
        
            Flog.writeline Espacios(Tabulador * 3) & "Tópico: " & rs_Event_Topic!topicheader & " - " & rs_Event_Topic!topicname
            
            StrSql = "SELECT * FROM topic_field "
            StrSql = StrSql & " WHERE topic_field.topicnro = " & rs_Event_Topic!topicnro
            StrSql = StrSql & " ORDER BY topic_field.tforden"
            OpenRecordset StrSql, rs_topic_field
            Do Until rs_topic_field.EOF
                
                StrSql = "SELECT * FROM event_topic_field "
                StrSql = StrSql & " WHERE event_topic_field.tfnro = " & rs_topic_field!tfnro
                StrSql = StrSql & " AND event_topic_field.eventnro = " & rs_Event_Topic!eventnro
                StrSql = StrSql & " AND event_topic_field.asoctype <> 3 " ' 3-Opcional no registrar
                OpenRecordset StrSql, rs_Event_Topic_Field
                
                Flog.Write Espacios(Tabulador * 4) & "Field: " & rs_topic_field!tfname & " - "
                
                If rs_Event_Topic_Field.EOF Then
                    ' LOG ARCHIVO
                    Flog.writeline "No hay campos para el topico y el evento que no sean opcionales."
                Else
                    ' Buscar el valor para el campo
                    programa = False
                    valorfijo = False
                    v_valor = ""
                    
                    ' Ver si es fijo o depende del evento
                    ' Ver si es mediante programa o valor fijo
                    
                    If rs_topic_field!valuetype = 3 Then ' Depende del evento
                        If rs_Event_Topic_Field!valuetype = 1 Then
                            ok_valor = True
                            valorfijo = True
                            v_valor = rs_Event_Topic_Field!fixedvalue
                            v_tipovalor = 1     ' Caracter - Valor no asoc. a entidad
                            Flog.writeline "El valor del campo para el topico y el evento depende del evento y es fijo. --> " & v_valor
                        Else
                            programa = True
                            v_prog = rs_Event_Topic_Field!valueprg
                            Flog.writeline "El valor del campo para el topico y el evento depende del evento y es por programa (" & v_prog & ")."
                        End If
                    ElseIf rs_topic_field!valuetype = 1 Then
                        ok_valor = True
                        valorfijo = True
                        v_valor = rs_topic_field!fixedvalue
                        v_tipovalor = 1         ' Caracter - Valor no asoc. a entidad
                        Flog.writeline "El valor del campo para el topico tiene valor fijo. --> " & v_valor
                    Else
                        programa = True
                        v_prog = rs_topic_field!valueprg
                        Flog.writeline "El valor del campo para el topico se busca mediante programa (" & v_prog & ")."
                    End If
                    
                    ' Obtener valor
                    If programa Then
                        Call obtener_valor(v_prog, intnro, 1, CLng(rs_Empleados!ternro), rs_event!eventnro, rs_Event_Topic!topicnro, rs_topic_field!tfnro, rs_Event_Topic_Field!asoctype, rs_event!eventcode, rs_Event_Topic!topicheader, rs_topic_field!tfname, v_tipovalor, v_valor, ok_valor, msg_valor)
                    End If
                        
                    ' CAMPO: event_topic_field.asoctype
                    ' 1-Obligatorio, 2-Opcional Registrar si se puede, 3-Opcional No Registrar
                    '
                    ' CAMPO: v_tipovalor
                    ' 1-Caracter, Valor no asoc. a entidad
                    ' 2-Valor formateado ya por el programa de obtencion
                    ' 3-Ya fue registrado por programa que obtuvo valor
                    ' 4-Se debe eliminar el topico
                    If v_tipovalor <> 3 And v_tipovalor <> 4 And Not ok_valor Then
                        ' Error, Warning
                        If rs_Event_Topic_Field!asoctype = 1 Then
                            If rs_Event_Topic_Field!asoctype = 1 Then
                                Call registro_log(intnro, 1, CLng(rs_Empleados!ternro), 2, rs_event!eventcode & ":" & rs_Event_Topic!topicheader + ":" + rs_topic_field!tfname & "(Obligatorio): " & v_prog)
                            Else
                                Call registro_log(intnro, 1, CLng(rs_Empleados!ternro), 2, rs_event!eventcode & ":" & rs_Event_Topic!topicheader + ":" + rs_topic_field!tfname & "(Opcional Registrar): " & v_prog)
                            End If
                        Else
                            If rs_Event_Topic_Field!asoctype = 1 Then
                                Call registro_log(intnro, 1, CLng(rs_Empleados!ternro), 1, rs_event!eventcode & ":" & rs_Event_Topic!topicheader + ":" + rs_topic_field!tfname & "(Obligatorio): " & v_prog)
                            Else
                                Call registro_log(intnro, 1, CLng(rs_Empleados!ternro), 1, rs_event!eventcode & ":" & rs_Event_Topic!topicheader + ":" + rs_topic_field!tfname & "(Opcional Registrar): " & v_prog)
                            End If
                        End If
                    
                        If rs_Event_Topic_Field!asoctype = 1 Then
                            GoTo bloque_empleado
                        End If
                    End If

                    ' Registrar valor para Int-Ter-Empresa-Topico-Campo
                    If v_valor <> "" And v_tipovalor <> 3 And v_tipovalor <> 4 Then
                        ' Grabar el registro
                        Flog.writeline Espacios(Tabulador * 4) & "  Registrar valor para Interfaz-Empleado-Empresa-Topico-Campo."
                        Call regeve01(intnro, 1, CLng(rs_Empleados!ternro), rs_event!eventnro, rs_Event_Topic!topicnro, rs_topic_field!tfnro, 0, v_valor, nuevo)
                    End If
                    
                    If v_tipovalor = 4 Then
                        ' Borrar los registros
                        Flog.writeline Espacios(Tabulador * 4) & "  Eliminar valor para Interfaz-Empleado-Empresa-Topico-Campo."
                        Call regeve02(intnro, 1, CLng(rs_Empleados!ternro), rs_event!eventnro, rs_Event_Topic!topicnro)
                        GoTo bloque_topico
                    End If
                End If
                    
                rs_Event_Topic_Field.Close
                    
                rs_topic_field.MoveNext
            Loop
            rs_topic_field.Close
            
bloque_topico:
            If rs_topic_field.State = adStateOpen Then rs_topic_field.Close
            If rs_Event_Topic_Field.State = adStateOpen Then rs_Event_Topic_Field.Close
                
            rs_Event_Topic.MoveNext
        Loop
            
        rs_Event_Topic.Close
                
bloque_evento:
        If rs_Event_Topic.State = adStateOpen Then rs_Event_Topic.Close
        If rs_topic_field.State = adStateOpen Then rs_topic_field.Close
        If rs_Event_Topic_Field.State = adStateOpen Then rs_Event_Topic_Field.Close
        
        rs_event.MoveNext
        
    Loop
        
    rs_event.Close
    
bloque_empleado:
    If rs_event.State = adStateOpen Then rs_event.Close
    If rs_Event_Topic.State = adStateOpen Then rs_Event_Topic.Close
    If rs_topic_field.State = adStateOpen Then rs_topic_field.Close
    If rs_Event_Topic_Field.State = adStateOpen Then rs_Event_Topic_Field.Close
    
    ' Registrar al empleado como generado
    ' Si no falto ningun campo obligatorio
    ' Si se le registro algo
    StrSql = "SELECT * FROM field_value "
    StrSql = StrSql & " WHERE intnro = " & intnro
    StrSql = StrSql & " AND empnro = 1 "
    StrSql = StrSql & " AND ternro = " & rs_Empleados!ternro
    OpenRecordset StrSql, rs_field_value
    If Not rs_field_value.EOF Then
        Call gendat01(intnro, 1, CLng(rs_Empleados!ternro))
    End If
    
    'Actualizo el progreso del Proceso
    Progreso = Progreso + IncPorc
    TiempoAcumulado = GetTickCount
    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & _
            ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & _
            "' WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
        
    ' Siguiente empleado
    rs_Empleados.MoveNext
    
Loop

'Fin de la transaccion
MyCommitTrans

Fin:

If rs_Empleados.State = adStateOpen Then rs_Empleados.Close
If rs_event.State = adStateOpen Then rs_event.Close
If rs_Event_Topic.State = adStateOpen Then rs_Event_Topic.Close
If rs_topic_field.State = adStateOpen Then rs_topic_field.Close
If rs_Event_Topic_Field.State = adStateOpen Then rs_Event_Topic_Field.Close

Set rs_Empleados = Nothing
Set rs_event = Nothing
Set rs_Event_Topic = Nothing
Set rs_topic_field = Nothing
Set rs_Event_Topic_Field = Nothing

Exit Sub

CE:
    Flog.writeline " ************************************************************************ "
    Flog.writeline " Error: " & Err.Description
    Flog.writeline " Ultima SQL Ejecutada: " & StrSql
    Flog.writeline " ************************************************************************ "
    HuboError = True
    MyRollbackTrans
Resume Next
    GoTo Fin
End Sub

Public Sub LevantarParamteros(ByVal bpronro As Long, ByVal parametros As String)
' --------------------------------------------------------------------------------------------
' Descripcion: Procedimiento para levantar los parametros pasados en batch_proceso en bprcparam
' Autor      : Fernando Favre
' Fecha      :
' Ult. Mod   :
' Fecha      :
' --------------------------------------------------------------------------------------------
Dim pos1 As Integer
Dim pos2 As Integer
Dim Separador As String

Dim intnro As Integer

    Separador = "@"
    ' Levanto cada parametro por separado
    If Not IsNull(parametros) Then
        If Len(parametros) >= 1 Then
            intnro = parametros
            
            'pos1 = 1
            'pos2 = InStr(pos1, parametros, Separador) - 1
            'fechadesde = Mid(parametros, pos1, pos2 - pos1 + 1)
            
            'pos1 = pos2 + 2
            'horahasta = Mid(parametros, pos1)
            
        Else
            Flog.writeline "El parametro informado por el proceso esta vacío."
        End If
    Else
        Flog.writeline "No se encontraron los parametros del proceso."
    End If
    
    Call Generacion(intnro)
End Sub

Private Sub evedet04(ByVal intnro As Integer, ByVal Empnro As Integer, ByVal ternro As Long, ByVal eventnro As Integer, ByRef tiene_evento As Boolean)
 Dim ant_pliqanio As Integer
 Dim ant_pliqmes As Integer
 Dim codetosend As String
 
 Dim rs_emp_attrib_value As New ADODB.Recordset
 Dim rs_periodo As New ADODB.Recordset
 Dim rs_acu_mes As New ADODB.Recordset
 Dim rs_b_acu_mes As New ADODB.Recordset

    tiene_evento = False
    
    ' Buscar si el empleado tiene asociado un valor con la entidad 15
    StrSql = "SELECT entity_value.codetosend FROM emp_attrib_value "
    StrSql = StrSql & " INNER JOIN entity_value ON emp_attrib_value.entvalnro = entity_value.entvalnro "
    StrSql = StrSql & " WHERE emp_attrib_value.ternro = " & ternro
    StrSql = StrSql & " AND entity_value.entnro = 15 "
'    StrSql = StrSql & " AND emp_attrib_value.empnro = " & empnro

    OpenRecordset StrSql, rs_emp_attrib_value
    
    If Not rs_emp_attrib_value.EOF Then
        codetosend = rs_emp_attrib_value!codetosend
        
        ' Buscar periodo de liq. Asociado.
        StrSql = "SELECT pliqanio, pliqmes  FROM interfaz "
        StrSql = StrSql & " INNER JOIN periodo ON interfaz.pliqnro = periodo.pliqnro "
        StrSql = StrSql & " WHERE interfaz.intnro = " & intnro
        
        OpenRecordset StrSql, rs_periodo
        
        If Not rs_periodo.EOF Then
            ant_pliqanio = CInt(rs_periodo!pliqanio)
            ant_pliqmes = CInt(rs_periodo!pliqmes) - 1
            If ant_pliqmes = 0 Then
                ant_pliqmes = 12
                ant_pliqanio = ant_pliqanio - 1
            End If
            
            StrSql = "SELECT acu_mes.acunro, acu_mes.ammonto, entity_value.aditfield3 FROM acu_mes "
            StrSql = StrSql & " INNER JOIN entity_map_value ON acu_mes.acunro = entity_map_value.domvalor AND entity_map_value.entmapnro = 6 "
            StrSql = StrSql & " INNER JOIN entity_value ON entity_map_value.imgvalor = entity_value.entvalnro "
            StrSql = StrSql & " WHERE "
            StrSql = StrSql & " acu_mes.ternro = " & ternro
            StrSql = StrSql & " AND acu_mes.amanio = " & rs_periodo!pliqanio
            StrSql = StrSql & " AND acu_mes.ammes = " & rs_periodo!pliqmes
            'StrSql = StrSql & " AND entity_map_value.empnro = " & empnro
            
            OpenRecordset StrSql, rs_acu_mes
            
            Do Until rs_acu_mes.EOF
                
                If InStr(1, rs_acu_mes!aditfield3, codetosend) > 0 Then
                    If rs_acu_mes!ammonto <> 0 Then
                        ' Busco el periodo anterior
                        StrSql = "SELECT acu_mes.acunro, acu_mes.ammonto FROM acu_mes "
                        StrSql = StrSql & " WHERE acu_mes.ternro = " & ternro
                        StrSql = StrSql & " AND acu_mes.amanio = " & ant_pliqanio
                        StrSql = StrSql & " AND acu_mes.ammes = " & ant_pliqmes
                        StrSql = StrSql & " AND acunro = " & rs_acu_mes!acunro
                        
                        OpenRecordset StrSql, rs_b_acu_mes
                        
                        If Not rs_b_acu_mes.EOF Then
                            If rs_acu_mes!ammonto <> rs_b_acu_mes!ammonto Then
                                tiene_evento = True
                            End If
                        Else
                            tiene_evento = True
                        End If
                         rs_b_acu_mes.Close
                        
                    End If
                End If
                rs_acu_mes.MoveNext
            Loop
            rs_acu_mes.Close
            
        End If
        
        rs_periodo.Close
        
    End If
    rs_emp_attrib_value.Close
    
End Sub

Private Sub evedet05(ByVal intnro As Integer, ByVal Empnro As Integer, ByVal ternro As Long, ByVal eventnro As Integer, ByRef tiene_evento As Boolean)
 Dim codetosend As String
 
 Dim rs_emp_attrib_value As New ADODB.Recordset
 Dim rs_periodo As New ADODB.Recordset
 Dim rs_acu_mes As New ADODB.Recordset

    tiene_evento = False
    
    ' Buscar si el empleado tiene asociado un valor con la entidad 15
    StrSql = "SELECT entity_value.codetosend FROM emp_attrib_value "
    StrSql = StrSql & " INNER JOIN entity_value ON emp_attrib_value.entvalnro = entity_value.entvalnro "
    StrSql = StrSql & " WHERE emp_attrib_value.ternro = " & ternro
    StrSql = StrSql & " AND entity_value.entnro = 15 "
'    StrSql = StrSql & " AND emp_attrib_value.empnro = " & empnro

    OpenRecordset StrSql, rs_emp_attrib_value
    
    If Not rs_emp_attrib_value.EOF Then
        codetosend = rs_emp_attrib_value!codetosend
        
        ' Buscar periodo de liq. Asociado.
        StrSql = "SELECT pliqanio, pliqmes  FROM interfaz "
        StrSql = StrSql & " INNER JOIN periodo ON interfaz.pliqnro = periodo.pliqnro "
        StrSql = StrSql & " WHERE interfaz.intnro = " & intnro
        
        OpenRecordset StrSql, rs_periodo
        
        If Not rs_periodo.EOF Then
            StrSql = "SELECT acu_mes.acunro, acu_mes.ammonto, entity_value.aditfield3 FROM acu_mes "
            StrSql = StrSql & " INNER JOIN entity_map_value ON acu_mes.acunro = entity_map_value.domvalor "
            StrSql = StrSql & " LEFT JOIN entity_value ON entity_map_value.imgvalor = entity_value.entvalnro "
            StrSql = StrSql & " WHERE entity_map_value.entmapnro = 6 "
            'StrSql = StrSql & " AND entity_map_value.empnro = " & empnro
            StrSql = StrSql & " AND acu_mes.ternro = " & ternro
            StrSql = StrSql & " AND acu_mes.amanio = " & rs_periodo!pliqanio
            StrSql = StrSql & " AND acu_mes.ammes = " & rs_periodo!pliqmes
            
            OpenRecordset StrSql, rs_acu_mes
            
            Do Until rs_acu_mes.EOF
                
                If InStr(1, rs_acu_mes!aditfield3, codetosend) > 0 Then
                    ' Buscar valor para el acumulador
                    If rs_acu_mes!ammonto <> 0 Then
                        tiene_evento = True
                    End If
                End If
                
                rs_acu_mes.MoveNext
            
            Loop
            rs_acu_mes.Close
            
        End If
        
        rs_periodo.Close
        
    End If
    rs_emp_attrib_value.Close

End Sub

Private Sub evedet06(ByVal intnro As Integer, ByVal Empnro As Integer, ByVal ternro As Long, ByVal eventnro As Integer, ByRef tiene_evento As Boolean)
 Dim seguir As Boolean
 Dim intfhasta As Date
 Dim mindiaslic As Integer
 Dim emp_licnro As Integer
 Dim elfechahasta As Date
 Dim elfechadesde As Date
 Dim tdnro As Integer
 Dim codetosend As String
 
 Dim rs_empleado As New ADODB.Recordset
 Dim rs_interfaz As New ADODB.Recordset
 Dim rs_intconfgen As New ADODB.Recordset
 Dim rs_emp_lic As New ADODB.Recordset
 Dim rs_intlicinfo As New ADODB.Recordset
 Dim rs_entity_map_value As New ADODB.Recordset
 Dim rs_event As New ADODB.Recordset
 Dim rs_b_Emp_Lic As New ADODB.Recordset
 
    seguir = True
    
    ' Si el empleado no esta activo no se tiene en cuenta para licencia
    StrSql = "SELECT empest FROM empleado "
    StrSql = StrSql & " WHERE empleado.ternro = " & ternro
    StrSql = StrSql & " AND empest = 0 "
    
    OpenRecordset StrSql, rs_empleado
    
    If Not rs_empleado.EOF Then
        seguir = False
    End If
    rs_empleado.Close
    
    
    ' Las licencias se deben informar solo si son por mas de 28 dias corridos, Hábiles y no hábiles
    ' Only those leaves with the same reason (ACTN_REASON) and a minimum length of four consecutive weeks
    ' before the interfacing date should be shown.
    
    If seguir Then
        ' Buscar el periodo de interfaz
        StrSql = "SELECT * FROM interfaz WHERE intnro = " & intnro
        OpenRecordset StrSql, rs_interfaz
        If rs_interfaz.EOF Then
            seguir = False
        Else
            intfhasta = rs_interfaz!intfhasta
        End If
        rs_interfaz.Close
    End If
    
    If seguir Then
        StrSql = "SELECT * FROM intconfgen"
        OpenRecordset StrSql, rs_intconfgen
        If rs_intconfgen.EOF Then
            seguir = False
        Else
            mindiaslic = rs_intconfgen!mindiaslic
        End If
        rs_intconfgen.Close
    End If
                
    If seguir Then
        ' Buscar licencia del empleado
        StrSql = "SELECT * FROM emp_lic "
        StrSql = StrSql & " WHERE emp_lic.empleado = " & ternro
        StrSql = StrSql & " AND elfechadesde <= " & ConvFecha(intfhasta) & " AND elfechahasta >= " & ConvFecha(intfhasta)
        OpenRecordset StrSql, rs_emp_lic
        If rs_emp_lic.EOF Then
            seguir = False
        Else
            emp_licnro = rs_emp_lic!emp_licnro
            elfechahasta = rs_emp_lic!elfechahasta
            elfechadesde = rs_emp_lic!elfechadesde
            tdnro = rs_emp_lic!tdnro
        End If
        rs_emp_lic.Close
    End If
        
    If seguir Then
        ' Verificar que no se haya informado la licencia ya
        StrSql = "SELECT * FROM intlicinfo "
        StrSql = StrSql & " WHERE emp_licnro = " & emp_licnro
        'StrSql = StrSql & " AND empnro = " & empnro
        OpenRecordset StrSql, rs_intlicinfo
        If Not rs_intlicinfo.EOF Then
            seguir = False
        End If
        rs_intlicinfo.Close
    End If
    
    If seguir Then
        ' Ver si es el tipo de licencia que le corresponde a este evento
        ' Nro. Mapeo, Valor del Dominio
        StrSql = "SELECT entity_value.codetosend FROM entity_map_value "
        StrSql = StrSql & " INNER JOIN entity_value ON entity_map_value.imgvalor = entity_value.entvalnro "
        StrSql = StrSql & " WHERE entity_map_value.domvalor = " & tdnro
        StrSql = StrSql & " AND entity_map_value.entmapnro = 10 "
        'StrSql = StrSql & " AND (entity_map.multiemp OR entity_map_value!empnro = " & empnro
        OpenRecordset StrSql, rs_entity_map_value
        If rs_entity_map_value.EOF Then
            seguir = False
        Else
            codetosend = rs_entity_map_value!codetosend
        End If
        rs_entity_map_value.Close
    End If
    
    If seguir Then
        StrSql = "SELECT * FROM event "
        StrSql = StrSql & " WHERE eventnro = " & eventnro
        OpenRecordset StrSql, rs_event
        If rs_event.EOF Then
            seguir = False
        Else
            If codetosend <> rs_event!eventcode Then
                seguir = False
            End If
        End If
        rs_event.Close
    End If
    
    If seguir Then
        If DateDiff("d", elfechadesde, elfechahasta) >= mindiaslic Then
            ' Registrar la licencia como informada
            StrSql = "INSERT INTO intlicinfo (empnro, emp_licnro, intnro) "
            StrSql = StrSql & " VALUES (1," & emp_licnro & ", " & intnro & ")"
            objConn.Execute StrSql, , adExecuteNoRecords
        Else
            ' Buscar otra licencia anterior que tenga las siguientes condiciones:
            '   1.- Debe ser de menos de 28 dias corridos
            '   2.- Debe ser del mismo tipo
            '   3.- Debe ser continua
            If Weekday(elfechadesde) = 2 Then
                ' Si elfechadesde es Lunes, debo considerar que la licencia termine un viernes, sabado o domingo
                StrSql = "SELECT * FROM emp_lic "
                StrSql = StrSql & " WHERE empleado = " & ternro
                StrSql = StrSql & " AND ((elfechahasta = " & DateAdd("d", -3, elfechadesde) ' viernes
                StrSql = StrSql & " ) OR (elfechahasta = " & DateAdd("d", -2, elfechadesde) ' sabado
                StrSql = StrSql & " ) OR (elfechahasta = " & DateAdd("d", -1, elfechadesde) ' domingo
                StrSql = StrSql & " )) AND tdnro = " & tdnro
            Else
                StrSql = "SELECT * FROM emp_lic "
                StrSql = StrSql & " WHERE empleado = " & ternro
                StrSql = StrSql & " AND (elfechahasta = " & DateAdd("d", -1, elfechadesde)
                StrSql = StrSql & " ) AND tdnro = " & tdnro
            End If
            
            OpenRecordset StrSql, rs_b_Emp_Lic
            
            If Not rs_b_Emp_Lic.EOF Then
                If ((DateDiff("d", rs_b_Emp_Lic!elfechadesde, rs_b_Emp_Lic!elfechahasta) < mindiaslic) And (DateDiff("d", rs_b_Emp_Lic!elfechadesde, elfechahasta) >= mindiaslic)) Then
                    ' Registrar la licencia como informada
                    StrSql = "INSERT INTO intlicinfo (empnro, emp_licnro, intnro) "
                    StrSql = StrSql & " VALUES (1," & rs_b_Emp_Lic!emp_licnro & ", " & intnro & ")"
                    objConn.Execute StrSql, , adExecuteNoRecords
                Else
                    seguir = False
                End If
            Else
                seguir = False
            End If
            
            rs_b_Emp_Lic.Close
            
        End If
    End If
    
    If seguir Then
        tiene_evento = True
    Else
        tiene_evento = False
    End If

End Sub

Private Sub evedet07(ByVal intnro As Integer, ByVal Empnro As Integer, ByVal ternro As Long, ByVal eventnro As Integer, ByRef tiene_evento As Boolean)
 Dim seguir As Boolean
 Dim intfhasta As Date
 Dim intfdesde As Date
 Dim emp_licnro As Integer
 
 Dim rs_interfaz As New ADODB.Recordset
 Dim rs_intconfgen As New ADODB.Recordset
 Dim rs_emp_lic As New ADODB.Recordset
 Dim rs_intlicinfo As New ADODB.Recordset
 
    seguir = True

    ' The event "Return after Leave" (RFL) is used to record the end of a leave
    ' previously notified by an event "Paid Leave" (PLA) or "Unpaid Leave" (LOA).

    ' Buscar una licencia de la persona que termine en este período de la interfaz

    ' Buscar el periodo de interfaz
    StrSql = "SELECT * FROM interfaz WHERE intnro = " & intnro
    OpenRecordset StrSql, rs_interfaz
    If rs_interfaz.EOF Then
        seguir = False
    Else
        intfhasta = rs_interfaz!intfhasta
        intfdesde = rs_interfaz!intfdesde
    End If
    rs_interfaz.Close
    
    If seguir Then
        StrSql = "SELECT * FROM intconfgen"
        OpenRecordset StrSql, rs_intconfgen
        If rs_intconfgen.EOF Then
            seguir = False
        End If
        rs_intconfgen.Close
    End If

    If seguir Then
        ' Buscar licencia del empleado
        StrSql = "SELECT * FROM emp_lic "
        StrSql = StrSql & " WHERE emp_lic.empleado = " & ternro
        StrSql = StrSql & " AND elfechahasta >= " & ConvFecha(DateAdd("d", -1, intfdesde)) & " AND elfechahasta <= " & ConvFecha(DateAdd("d", -1, intfhasta))
        OpenRecordset StrSql, rs_emp_lic
        If rs_emp_lic.EOF Then
            seguir = False
        Else
            emp_licnro = rs_emp_lic!emp_licnro
        End If
        rs_emp_lic.Close
    End If
        
    If seguir Then
        ' Verificar si dicha licencia fue informada
        StrSql = "SELECT * FROM intlicinfo "
        StrSql = StrSql & " WHERE emp_licnro = " & emp_licnro
        'StrSql = StrSql & " AND empnro = " & empnro
        OpenRecordset StrSql, rs_intlicinfo
        If rs_intlicinfo.EOF Then
            seguir = False
        End If
        rs_intlicinfo.Close
    End If

    If seguir Then
        tiene_evento = True
    Else
        tiene_evento = False
    End If

End Sub

Private Sub evedet08(ByVal intnro As Integer, ByVal Empnro As Integer, ByVal ternro As Long, ByVal eventnro As Integer, ByRef tiene_evento As Boolean)
 Dim seguir As Boolean
 Dim codetosend1 As String
 Dim intfhasta As Date
 Dim intfdesde As Date
 Dim eventcode As String
 Dim codetosend2 As String
 Dim caunro As Integer
  
 Dim rs_empleado As New ADODB.Recordset
 Dim rs_emp_attrib_value As New ADODB.Recordset
 Dim rs_interfaz As New ADODB.Recordset
 Dim rs_fases As New ADODB.Recordset
 Dim rs_entity_map_value As New ADODB.Recordset
 Dim rs_event As New ADODB.Recordset

    seguir = True
    
    ' Si el empleado esta activo no se tiene en cuenta
    StrSql = "SELECT empest FROM empleado "
    StrSql = StrSql & " WHERE empleado.ternro = " & ternro
    StrSql = StrSql & " AND empest = -1"
    
    OpenRecordset StrSql, rs_empleado
    
    If Not rs_empleado.EOF Then
        seguir = False
    End If
    rs_empleado.Close
    
    
    ' The event "Termination" (TER) reflects the employee's definitive departure from the BNP Paribas Group, for a reason other than retirement.
    ' The employees concerned by this event are local and dual employees with, therefore, an impact on record numbers 0 and 2 respectively.
    
    ' This event is not used when the employee leaves the company for any of the following reasons :
    ' ú   Expatriation or secondment or on loan,
    ' ú   retirement.
    
    ' An employee record with a Termination (TER)  event should no longer receive other events (except in the case of "re-hiring").
    
    ' The effective date of the event "termination" (TER) is the first day the former employee is no more working for the company.
    
    If seguir Then
        ' Buscar si el empleado tiene asociado un valor con la entidad 15
        StrSql = "SELECT entity_value.codetosend FROM emp_attrib_value "
        StrSql = StrSql & " INNER JOIN entity_value ON emp_attrib_value.entvalnro = entity_value.entvalnro "
        StrSql = StrSql & " WHERE emp_attrib_value.ternro = " & ternro
        StrSql = StrSql & " AND entity_value.entnro = 15 "
    '    StrSql = StrSql & " AND emp_attrib_value.empnro = " & empnro
    
        OpenRecordset StrSql, rs_emp_attrib_value
        
        If rs_emp_attrib_value.EOF Then
            seguir = False
        Else
            codetosend1 = rs_emp_attrib_value!codetosend
        End If
        rs_emp_attrib_value.Close
    End If
    
    If seguir Then
        ' Verificar que el Registro no sea 0 o 2
        If codetosend1 <> "0" And codetosend1 <> "2" Then
            seguir = False
        End If
    End If

    If seguir Then
        ' Buscar el periodo de interfaz
        StrSql = "SELECT * FROM interfaz WHERE intnro = " & intnro
        OpenRecordset StrSql, rs_interfaz
        If rs_interfaz.EOF Then
            seguir = False
        Else
            intfhasta = rs_interfaz!intfhasta
            intfdesde = rs_interfaz!intfdesde
        End If
        rs_interfaz.Close
    End If

    If seguir Then
        ' Buscar Baja dentro del periodo de la interfaz
        StrSql = "SELECT * FROM fases "
        StrSql = StrSql & " WHERE empleado = " & ternro                                                         ' de la persona en cuestion
        StrSql = StrSql & " AND bajfec >=" & ConvFecha(intfdesde) & " AND bajfec <= " & ConvFecha(intfhasta)    ' dentro del periodo de la interfaz
        StrSql = StrSql & " AND estado = 0 "                                                                    ' Baja
        StrSql = StrSql & " ORDER BY altfec DESC"
        OpenRecordset StrSql, rs_fases
        If rs_fases.EOF Then
            seguir = False
        Else
            caunro = rs_fases!caunro
        End If
        rs_fases.Close
    End If

    If seguir Then
        ' Verificar si es el tipo de baja de este evento */
        ' Nro. Mapeo, Valor del Dominio */
        ' Determino Si existe algun mapeo */
        StrSql = "SELECT entity_value.codetosend FROM entity_map_value "
        StrSql = StrSql & " LEFT JOIN entity_value ON entity_map_value.imgvalor = entity_value.entvalnro "
        StrSql = StrSql & " WHERE entity_map_value.domvalor = " & caunro
        StrSql = StrSql & " AND entity_map_value.entmapnro = 12 "
        'StrSql = StrSql & " AND entity_map_value.empnro = " & empnro
        OpenRecordset StrSql, rs_entity_map_value
        If rs_entity_map_value.EOF Then
            seguir = False
        Else
            codetosend1 = rs_entity_map_value!codetosend
        End If
        rs_entity_map_value.Close
    End If

    If seguir Then
        ' Buscar descripción del Evento
        StrSql = "SELECT * FROM event WHERE eventnro = " & eventnro
        OpenRecordset StrSql, rs_event
        If rs_event.EOF Then
            seguir = False
        Else
            eventcode = rs_event!eventcode
        End If
        rs_event.Close
    End If

    If seguir Then
        If eventcode <> codetosend1 Then
            seguir = False
        End If
    End If
    
    If seguir Then
        tiene_evento = True
    Else
        tiene_evento = False
    End If
End Sub

Private Sub evedet09(ByVal intnro As Integer, ByVal Empnro As Integer, ByVal ternro As Long, ByVal eventnro As Integer, ByRef tiene_evento As Boolean)
 Dim seguir As Boolean
 Dim codetosend1 As String
 Dim intfhasta As Date
 Dim intfdesde As Date
 Dim eventcode As String
 Dim codetosend2 As String
 Dim caunro As Integer
  
 Dim rs_empleado As New ADODB.Recordset
 Dim rs_emp_attrib_value As New ADODB.Recordset
 Dim rs_interfaz As New ADODB.Recordset
 Dim rs_fases As New ADODB.Recordset
 Dim rs_entity_map_value As New ADODB.Recordset
 Dim rs_event As New ADODB.Recordset

    seguir = True
    
    ' Si el empleado esta activo no se tiene en cuenta
    StrSql = "SELECT empest FROM empleado "
    StrSql = StrSql & " WHERE empleado.ternro = " & ternro
    StrSql = StrSql & " AND empest = -1"
    
    OpenRecordset StrSql, rs_empleado
    
    If Not rs_empleado.EOF Then
        seguir = False
    End If
    rs_empleado.Close
    
    
    ' The event "Termination" (TER) reflects the employee's definitive departure from the BNP Paribas Group, for a reason other than retirement.
    ' The employees concerned by this event are local and dual employees with, therefore, an impact on record numbers 0 and 2 respectively.
    
    ' This event is not used when the employee leaves the company for any of the following reasons :
    ' ú   Expatriation or secondment or on loan,
    ' ú   retirement.
    
    ' An employee record with a Termination (TER)  event should no longer receive other events (except in the case of "re-hiring").
    
    ' The effective date of the event "termination" (TER) is the first day the former employee is no more working for the company.
    
    If seguir Then
        ' Buscar si el empleado tiene asociado un valor con la entidad 15
        StrSql = "SELECT entity_value.codetosend FROM emp_attrib_value "
        StrSql = StrSql & " INNER JOIN entity_value ON emp_attrib_value.entvalnro = entity_value.entvalnro "
        StrSql = StrSql & " WHERE emp_attrib_value.ternro = " & ternro
        StrSql = StrSql & " AND entity_value.entnro = 15 "
    '    StrSql = StrSql & " AND emp_attrib_value.empnro = " & empnro
    
        OpenRecordset StrSql, rs_emp_attrib_value
        
        If rs_emp_attrib_value.EOF Then
            seguir = False
        Else
            codetosend1 = rs_emp_attrib_value!codetosend
        End If
        rs_emp_attrib_value.Close
    End If
    
    If seguir Then
        ' Verificar que el Registro no sea 1 o 3
        If codetosend1 <> "1" And codetosend1 <> "3" Then
            seguir = False
        End If
    End If

    If seguir Then
        ' Buscar el periodo de interfaz
        StrSql = "SELECT * FROM interfaz WHERE intnro = " & intnro
        OpenRecordset StrSql, rs_interfaz
        If rs_interfaz.EOF Then
            seguir = False
        Else
            intfhasta = rs_interfaz!intfhasta
            intfdesde = rs_interfaz!intfdesde
        End If
        rs_interfaz.Close
    End If

    If seguir Then
        ' Buscar Baja dentro del periodo de la interfaz
        StrSql = "SELECT * FROM fases "
        StrSql = StrSql & " WHERE empleado = " & ternro                                                         ' de la persona en cuestion
        StrSql = StrSql & " AND bajfec >=" & ConvFecha(intfdesde) & " AND bajfec <= " & ConvFecha(intfhasta)    ' dentro del periodo de la interfaz
        StrSql = StrSql & " AND estado = 0 "                                                                    ' Baja
        StrSql = StrSql & " ORDER BY altfec DESC"
        OpenRecordset StrSql, rs_fases
        If rs_fases.EOF Then
            seguir = False
        Else
            caunro = rs_fases!caunro
        End If
        rs_fases.Close
    End If

    If seguir Then
        ' Verificar si es el tipo de baja de este evento */
        ' Nro. Mapeo, Valor del Dominio */
        ' Determino Si existe algun mapeo */
        StrSql = "SELECT entity_value.codetosend FROM entity_map_value "
        StrSql = StrSql & " LEFT JOIN entity_value ON entity_map_value.imgvalor = entity_value.entvalnro "
        StrSql = StrSql & " WHERE entity_map_value.domvalor = " & caunro
        StrSql = StrSql & " AND entity_map_value.entmapnro = 12 "
        'StrSql = StrSql & " AND entity_map_value.empnro = " & empnro
        OpenRecordset StrSql, rs_entity_map_value
        If rs_entity_map_value.EOF Then
            seguir = False
        Else
            codetosend1 = rs_entity_map_value!codetosend
        End If
        rs_entity_map_value.Close
    End If

    If seguir Then
        ' Buscar descripción del Evento
        StrSql = "SELECT * FROM event WHERE eventnro = " & eventnro
        OpenRecordset StrSql, rs_event
        If rs_event.EOF Then
            seguir = False
        Else
            eventcode = rs_event!eventcode
        End If
        rs_event.Close
    End If

    If seguir Then
        If eventcode <> codetosend1 Then
            seguir = False
        End If
    End If
    
    If seguir Then
        tiene_evento = True
    Else
        tiene_evento = False
    End If
End Sub

Private Sub evedet10(ByVal intnro As Integer, ByVal Empnro As Integer, ByVal ternro As Long, ByVal eventnro As Integer, ByRef tiene_evento As Boolean)
 Dim seguir As Boolean
 Dim codetosend1 As String
 Dim intfhasta As Date
 Dim intfdesde As Date
 Dim eventcode As String
 Dim codetosend2 As String
 Dim caunro As Integer
 
 Dim rs_empleado As New ADODB.Recordset
 Dim rs_emp_attrib_value As New ADODB.Recordset
 Dim rs_interfaz As New ADODB.Recordset
 Dim rs_fases As New ADODB.Recordset
 Dim rs_entity_map_value As New ADODB.Recordset
 Dim rs_event As New ADODB.Recordset

    seguir = True
    
    ' Si el empleado esta activo no se tiene en cuenta
    StrSql = "SELECT empest FROM empleado "
    StrSql = StrSql & " WHERE empleado.ternro = " & ternro
    StrSql = StrSql & " AND empest = -1"
    
    OpenRecordset StrSql, rs_empleado
    
    If Not rs_empleado.EOF Then
        seguir = False
    End If
    rs_empleado.Close
    
    ' The event "Global Assignment - Home Company" (END) is used by the Home Company
    ' (employee record number 0) to record the departure of an employee being seconded, expatriated or on loan.
    ' This event impacts FTE, employee class and, optionally, the employee's fixed compensation.
    
    If seguir Then
        ' Buscar si el empleado tiene asociado un valor con la entidad 15
        StrSql = "SELECT entity_value.codetosend FROM emp_attrib_value "
        StrSql = StrSql & " INNER JOIN entity_value ON emp_attrib_value.entvalnro = entity_value.entvalnro "
        StrSql = StrSql & " WHERE emp_attrib_value.ternro = " & ternro
        StrSql = StrSql & " AND entity_value.entnro = 15 "
    '    StrSql = StrSql & " AND emp_attrib_value.empnro = " & empnro
    
        OpenRecordset StrSql, rs_emp_attrib_value
        
        If rs_emp_attrib_value.EOF Then
            seguir = False
        Else
            codetosend1 = rs_emp_attrib_value!codetosend
        End If
        rs_emp_attrib_value.Close
    End If
    
    If seguir Then
        ' Verificar que el Registro no sea 0
        If codetosend1 <> "0" Then
            seguir = False
        End If
    End If

    If seguir Then
        ' Buscar el periodo de interfaz
        StrSql = "SELECT * FROM interfaz WHERE intnro = " & intnro
        OpenRecordset StrSql, rs_interfaz
        If rs_interfaz.EOF Then
            seguir = False
        Else
            intfhasta = rs_interfaz!intfhasta
            intfdesde = rs_interfaz!intfdesde
        End If
        rs_interfaz.Close
    End If

    If seguir Then
        ' Buscar Baja dentro del periodo de la interfaz
        StrSql = "SELECT * FROM fases "
        StrSql = StrSql & " WHERE empleado = " & ternro                                                         ' de la persona en cuestion
        StrSql = StrSql & " AND bajfec >=" & ConvFecha(intfdesde) & " AND bajfec <= " & ConvFecha(intfhasta)    ' dentro del periodo de la interfaz
        StrSql = StrSql & " AND estado = 0 "                                                                    ' Baja
        StrSql = StrSql & " ORDER BY altfec DESC"
        OpenRecordset StrSql, rs_fases
        If rs_fases.EOF Then
            seguir = False
        Else
            caunro = rs_fases!caunro
        End If
        rs_fases.Close
    End If

    If seguir Then
        ' Verificar si es el tipo de baja de este evento */
        ' Nro. Mapeo, Valor del Dominio */
        ' Determino Si existe algun mapeo */
        StrSql = "SELECT entity_value.codetosend FROM entity_map_value "
        StrSql = StrSql & " LEFT JOIN entity_value ON entity_map_value.imgvalor = entity_value.entvalnro "
        StrSql = StrSql & " WHERE entity_map_value.domvalor = " & caunro
        StrSql = StrSql & " AND entity_map_value.entmapnro = 12 "
        'StrSql = StrSql & " AND entity_map_value.empnro = " & empnro
        OpenRecordset StrSql, rs_entity_map_value
        If rs_entity_map_value.EOF Then
            seguir = False
        Else
            codetosend1 = rs_entity_map_value!codetosend
        End If
        rs_entity_map_value.Close
    End If

    If seguir Then
        ' Buscar descripción del Evento
        StrSql = "SELECT * FROM event WHERE eventnro = " & eventnro
        OpenRecordset StrSql, rs_event
        If rs_event.EOF Then
            seguir = False
        Else
            eventcode = rs_event!eventcode
        End If
        rs_event.Close
    End If

    If seguir Then
        If eventcode <> codetosend1 Then
            seguir = False
        End If
    End If
    
    If seguir Then
        tiene_evento = True
    Else
        tiene_evento = False
    End If
End Sub

Private Sub evedet11(ByVal intnro As Integer, ByVal Empnro As Integer, ByVal ternro As Long, ByVal eventnro As Integer, ByRef tiene_evento As Boolean)
 Dim seguir As Boolean
 Dim sysid As String
 Dim empleg As Long
 
 Dim rs_empleado As New ADODB.Recordset
 Dim rs_intconfgen As New ADODB.Recordset
 Dim rs_field_value As New ADODB.Recordset
 
    seguir = True
    tiene_evento = False
    
    ' Buscar legajo del empleado
    StrSql = "SELECT empleg FROM empleado "
    StrSql = StrSql & " WHERE empleado.ternro = " & ternro
    
    OpenRecordset StrSql, rs_empleado
    
    If rs_empleado.EOF Then
        seguir = False
    Else
        empleg = rs_empleado!empleg
    End If
    rs_empleado.Close
    
    If seguir Then
        StrSql = "SELECT valor FROM field_value "
        StrSql = StrSql & " WHERE archivado = -1 "
        'StrSql = StrSql & " AND empnro = " & empnro
        StrSql = StrSql & " AND intnro <> " & intnro
        StrSql = StrSql & " AND ternro = " & ternro
        ' Eventos DTC_XXX o DTA_PER anteriores
        StrSql = StrSql & " AND (eventnro = 22 "
        StrSql = StrSql & " OR eventnro = 24"
        StrSql = StrSql & " OR eventnro = 25"
        StrSql = StrSql & " OR eventnro = 26)"
        StrSql = StrSql & " AND tfnro = 352"    ' LOCAL_EMPLID
        StrSql = StrSql & " AND topicnro = 5"   ' MAT
        StrSql = StrSql & " ORDER BY intnro DESC "
        OpenRecordset StrSql, rs_field_value
        Do Until rs_field_value.EOF Or tiene_evento
            If CStr(rs_field_value!Valor) <> CStr(empleg) Then
                tiene_evento = True
            End If
            rs_field_value.MoveNext
        Loop
        rs_field_value.Close
    End If

    If seguir And Not tiene_evento Then
        ' Buscar ident. del sistema
        StrSql = "SELECT sysid FROM intconfgen"
        OpenRecordset StrSql, rs_intconfgen
        If rs_intconfgen.EOF Then
            seguir = False
        Else
            sysid = rs_intconfgen!sysid
        End If
        rs_intconfgen.Close
    End If
    
    If seguir And Not tiene_evento Then
        StrSql = "SELECT valor FROM field_value "
        StrSql = StrSql & " WHERE archivado = -1 "
        'StrSql = StrSql & " AND empnro = " & empnro
        StrSql = StrSql & " AND intnro <> " & intnro
        StrSql = StrSql & " AND ternro = " & ternro
        ' Eventos DTC_XXX o DTA_PER anteriores
        StrSql = StrSql & " AND (eventnro = 22 "
        StrSql = StrSql & " OR eventnro = 24"
        StrSql = StrSql & " OR eventnro = 25"
        StrSql = StrSql & " OR eventnro = 26)"
        StrSql = StrSql & " AND tfnro = 351"    ' SYS_ID
        StrSql = StrSql & " AND topicnro = 5"   ' MAT
        StrSql = StrSql & " ORDER BY intnro DESC "
        OpenRecordset StrSql, rs_field_value
        Do Until rs_field_value.EOF Or tiene_evento
            If CStr(rs_field_value!Valor) <> CStr(sysid) Then
                tiene_evento = True
            End If
            rs_field_value.MoveNext
        Loop
        rs_field_value.Close
    End If
End Sub

Private Sub evedet12(ByVal intnro As Integer, ByVal Empnro As Integer, ByVal ternro As Long, ByVal eventnro As Integer, ByRef tiene_evento As Boolean)
 Dim seguir As Boolean
 Dim intfhasta As Date
 Dim intfdesde As Date
 Dim inthhasta As Long
 Dim inthdesde As Long
 Dim Valor As String
 Dim hora
 Dim aud_hor As Long
 Dim aud_fec As Date
 Dim caudnro As Integer
 Dim aud_campnro As Integer
 Dim acnro As Integer
 
 Dim rs_interfaz As New ADODB.Recordset
 Dim rs_intauditoria As New ADODB.Recordset
 Dim rs_auditoria As New ADODB.Recordset
 Dim rs_auditoria_prev As New ADODB.Recordset

    seguir = True
    tiene_evento = False
    
    ' Buscar el periodo de interfaz
    StrSql = "SELECT * FROM interfaz WHERE intnro = " & intnro
    OpenRecordset StrSql, rs_interfaz
    If rs_interfaz.EOF Then
        seguir = False
    Else
        intfhasta = rs_interfaz!intfhasta
        intfdesde = rs_interfaz!intfdesde
        inthhasta = CLng(rs_interfaz!inthhasta)
        inthdesde = CLng(rs_interfaz!inthdesde)
    End If
    rs_interfaz.Close
    
    
    If seguir Then
        StrSql = "SELECT * FROM intauditoria "
        StrSql = StrSql & " INNER JOIN event_topic_field ON intauditoria.tfnro = event_topic_field.tfnro "
        StrSql = StrSql & " WHERE event_topic_field.asoctype = 1 OR event_topic_field.asoctype = 2 "
        
        OpenRecordset StrSql, rs_intauditoria
        
        Do Until rs_intauditoria.EOF Or tiene_evento
            caudnro = rs_intauditoria!caudnro
            acnro = rs_intauditoria!acnro
            aud_campnro = rs_intauditoria!aud_campnro
             
            StrSql = "SELECT aud_fec, aud_hor, aud_actual, audnro FROM auditoria "
            StrSql = StrSql & " WHERE auditoria.aud_ternro = " & ternro
            'StrSql = StrSql & " AND auditoria.aud_emp = " & empnro
            StrSql = StrSql & " AND auditoria.caudnro = " & caudnro
            StrSql = StrSql & " AND auditoria.aud_campnro = " & aud_campnro
            StrSql = StrSql & " AND auditoria.acnro = " & acnro
            StrSql = StrSql & " AND auditoria.aud_fec  >= " & ConvFecha(intfdesde) & " AND auditoria.aud_fec <= " & ConvFecha(intfhasta)
            StrSql = StrSql & " ORDER BY auditoria.aud_fec DESC"
            
            OpenRecordset StrSql, rs_auditoria
            
            If Not rs_auditoria.EOF Then
                hora = Split(rs_auditoria!aud_hor, ":")
                aud_hor = CLng(hora(0))
                aud_hor = aud_hor * CLng(100) + CLng(hora(1))
                aud_hor = aud_hor * CLng(100) + CLng(hora(2))
                aud_fec = rs_auditoria!aud_fec
                If (aud_fec >= intfdesde And aud_fec <= intfhasta) Or (aud_fec = intfdesde And aud_hor >= inthdesde) Or (aud_fec = intfhasta And aud_hor <= inthhasta) Then
                    If EsNulo(rs_auditoria!aud_actual) Then
                        Valor = ""
                    Else
                        Valor = rs_auditoria!aud_actual
                    End If
        
                    ' Verificar que el campo ya no halla tenido este valor en el mismo período,
                    ' si esto sucediera, el campo nunca cambió de valor
                    
                    StrSql = "SELECT * FROM auditoria "
                    StrSql = StrSql & " WHERE auditoria.aud_ternro = " & ternro
                    'StrSql = StrSql & " AND auditoria.aud_emp = " & empnro
                    StrSql = StrSql & " AND auditoria.caudnro = " & rs_intauditoria!caudnro
                    StrSql = StrSql & " AND auditoria.aud_campnro = " & aud_campnro
                    StrSql = StrSql & " AND auditoria.acnro = " & rs_intauditoria!acnro
                    StrSql = StrSql & " AND auditoria.aud_fec >= " & ConvFecha(intfdesde) & " AND auditoria.aud_fec <= " & ConvFecha(intfhasta)
                    StrSql = StrSql & " AND auditoria.aud_ant = '" & Valor & "'"
                    StrSql = StrSql & " AND auditoria.audnro <> " & rs_auditoria!audnro
                    StrSql = StrSql & " ORDER BY auditoria.aud_fec DESC"
                    OpenRecordset StrSql, rs_auditoria_prev
                    If rs_auditoria_prev.EOF Then
                        tiene_evento = True
                    Else
                        hora = Split(rs_auditoria_prev!aud_hor, ":")
                        aud_hor = CLng(hora(0))
                        aud_hor = aud_hor * CLng(100) + CLng(hora(1))
                        aud_hor = aud_hor * CLng(100) + CLng(hora(2))
                        aud_fec = rs_auditoria!aud_fec
                        If Not ((aud_fec >= intfdesde And aud_fec <= intfhasta) Or (aud_fec = intfdesde And aud_hor >= inthdesde) Or (aud_fec = intfhasta And aud_hor <= inthhasta)) Then
                            tiene_evento = True
                        End If
                    End If
                    
                    rs_auditoria_prev.Close
                    
                End If
            End If
            
            rs_auditoria.Close
            
            rs_intauditoria.MoveNext
        Loop
        
        rs_intauditoria.Close
        
    End If

End Sub

Private Sub evedet13(ByVal intnro As Integer, ByVal Empnro As Integer, ByVal ternro As Long, ByVal eventnro As Integer, ByRef tiene_evento As Boolean)
 Dim seguir As Boolean
 Dim intfhasta As Date
 Dim intfdesde As Date
 Dim inthhasta As Long
 Dim inthdesde As Long
 Dim hora
 Dim aud_hor As Long
 Dim aud_fec As Date
 Dim eventcode As String
 
 Dim rs_interfaz As New ADODB.Recordset
 Dim rs_eveauditoria As New ADODB.Recordset
 Dim rs_auditoria As New ADODB.Recordset
 Dim rs_event As New ADODB.Recordset
 Dim rs_entity_value As New ADODB.Recordset
 Dim rs_tip_grup_ter As New ADODB.Recordset
 
    seguir = True
    tiene_evento = True
    
    ' Buscar el periodo de interfaz
    StrSql = "SELECT * FROM interfaz WHERE intnro = " & intnro
    OpenRecordset StrSql, rs_interfaz
    If rs_interfaz.EOF Then
        seguir = False
        tiene_evento = False
    Else
        intfhasta = rs_interfaz!intfhasta
        intfdesde = rs_interfaz!intfdesde
        inthhasta = CLng(rs_interfaz!inthhasta)
        inthdesde = CLng(rs_interfaz!inthdesde)
    End If
    rs_interfaz.Close
    
    If seguir Then
        StrSql = "SELECT * FROM eveauditoria "
        StrSql = StrSql & " WHERE eveauditoria.eventnro = " & eventnro
        StrSql = StrSql & " ORDER BY orden"
        OpenRecordset StrSql, rs_eveauditoria
        Do Until rs_eveauditoria.EOF
            StrSql = "SELECT aud_fec, aud_hor, aud_actual, audnro FROM auditoria "
            StrSql = StrSql & " WHERE auditoria.aud_ternro = " & ternro
            'StrSql = StrSql & " AND auditoria.aud_emp = " & empnro
            StrSql = StrSql & " AND auditoria.caudnro = " & rs_eveauditoria!caudnro
            StrSql = StrSql & " AND auditoria.aud_campnro = " & rs_eveauditoria!aud_campnro
            StrSql = StrSql & " AND auditoria.acnro = " & rs_eveauditoria!acnro
            StrSql = StrSql & " AND auditoria.aud_fec >= " & ConvFecha(intfdesde) & " AND auditoria.aud_fec <= " & ConvFecha(intfhasta)
            StrSql = StrSql & " ORDER BY auditoria.aud_fec DESC"
            
            OpenRecordset StrSql, rs_auditoria
            
            If rs_auditoria.EOF Then
                tiene_evento = False
            Else
                hora = Split(rs_auditoria!aud_hor, ":")
                aud_hor = CLng(hora(0))
                aud_hor = aud_hor * CLng(100) + CLng(hora(1))
                aud_hor = aud_hor * CLng(100) + CLng(hora(2))
                aud_fec = rs_auditoria!aud_fec
                If Not ((aud_fec >= intfdesde And aud_fec <= intfhasta) Or (aud_fec = intfdesde And aud_hor >= inthdesde) Or (aud_fec = intfhasta And aud_hor <= inthhasta)) Then
                    tiene_evento = False
                End If
            End If
            
            rs_auditoria.Close
            
            rs_eveauditoria.MoveNext
        
        Loop
        
        rs_eveauditoria.Close
        
    End If
    
    If tiene_evento Then
        ' Buscar Grupo correspondiente a este evento
        StrSql = "SELECT * FROM event WHERE event.eventnro = " & eventnro
        OpenRecordset StrSql, rs_event
        If rs_event.EOF Then
            tiene_evento = False
        Else
            eventcode = rs_event!eventcode
        End If
        rs_event.Close
    End If
    
    If tiene_evento Then
        StrSql = "SELECT entity_value.codetosend, entity_map_value.domvalor FROM entity_value "
        StrSql = StrSql & " INNER JOIN entity_map_value ON entity_map_value.imgvalor = entity_value.entvalnro "
        StrSql = StrSql & " WHERE entity_map_value.entmapnro = 15 "
        'StrSql = StrSql & " AND entity_map_value.empnro = " & empnro
        StrSql = StrSql & " AND entity_value.codetosend = '" & eventcode & "'"
        OpenRecordset StrSql, rs_entity_value
        If rs_entity_value.EOF Then
            tiene_evento = False
        Else
            ' Verificar si el empleado esta en el grupo
            StrSql = "SELECT * FROM his_estructura "
            StrSql = StrSql & " WHERE his_estructura.estrnro = " & rs_entity_value!domvalor
            StrSql = StrSql & " AND his_estructura.ternro = " & ternro
            StrSql = StrSql & " AND his_estructura.htethasta is null"
            OpenRecordset StrSql, rs_tip_grup_ter
            If rs_tip_grup_ter.EOF Then
                tiene_evento = False
            End If
            rs_tip_grup_ter.Close
        End If
        rs_entity_value.Close
    End If

End Sub

Private Sub evedet14(ByVal intnro As Integer, ByVal Empnro As Integer, ByVal ternro As Long, ByVal eventnro As Integer, ByRef tiene_evento As Boolean)
 Dim seguir As Boolean
 Dim intfhasta As Date
 Dim intfdesde As Date
 Dim inthhasta As Long
 Dim inthdesde As Long
 Dim claveult As Long
 Dim clavepri As Long
 Dim actult As String
 Dim antpri As String
 
 Dim rs_interfaz As New ADODB.Recordset
 Dim rs_entity_value As New ADODB.Recordset
 Dim rs_auditoria As New ADODB.Recordset

    seguir = True
    
    ' Buscar el periodo de interfaz
    StrSql = "SELECT * FROM interfaz WHERE intnro = " & intnro
    OpenRecordset StrSql, rs_interfaz
    If rs_interfaz.EOF Then
        seguir = False
    Else
        intfhasta = rs_interfaz!intfhasta
        intfdesde = rs_interfaz!intfdesde
        inthhasta = CLng(rs_interfaz!inthhasta)
        inthdesde = CLng(rs_interfaz!inthdesde)
    End If
    rs_interfaz.Close
    
    If seguir Then
        ' Buscar Eventos no automaticos para el empleado
        StrSql = "SELECT * FROM emp_attrib_value "
        StrSql = StrSql & " INNER JOIN entity_value ON emp_attrib_value.entvalnro = entity_value.entvalnro "
        StrSql = StrSql & " WHERE emp_attrib_value.ternro = " & ternro
        'StrSql = StrSql & " AND emp_attrib_value.empnro = " & empnro
        StrSql = StrSql & " AND emp_attrib_value.fecha >= " & ConvFecha(intfdesde)
        StrSql = StrSql & " AND emp_attrib_value.fecha <= " & ConvFecha(intfhasta)
        StrSql = StrSql & " AND (entity_value.entnro = 2 "  'BUSSINES_UNIT
        StrSql = StrSql & " OR entity_value.entnro = 3 "    'DEPTIP
        StrSql = StrSql & " OR entity_value.entnro = 19)"   'LOCATION
        
        OpenRecordset StrSql, rs_entity_value
        
        If rs_entity_value.EOF Then
            seguir = False
        End If
        
        rs_entity_value.Close
        
    End If
    
    If Not seguir Then
        ' Cambio de sucursal - LOCATION - Buscar en la auditoria
        StrSql = "SELECT * FROM auditoria "
        StrSql = StrSql & " WHERE auditoria.aud_ternro = " & ternro
        'StrSql = StrSql & " AND auditoria.aud_emp = " & empnro
        StrSql = StrSql & " AND auditoria.caudnro = 34 "
        StrSql = StrSql & " AND auditoria.aud_campnro = 89 "
        StrSql = StrSql & " AND auditoria.acnro = 2 "
        StrSql = StrSql & " AND auditoria.aud_fec >= " & ConvFecha(intfdesde)
        StrSql = StrSql & " AND auditoria.aud_fec <= " & ConvFecha(intfhasta)
        StrSql = StrSql & " ORDER BY audnro"
        OpenRecordset StrSql, rs_auditoria
        If Not rs_auditoria.EOF Then
            rs_auditoria.MoveFirst
            actult = rs_auditoria!aud_actual
            claveult = CLng(rs_auditoria!audnro)
            
            rs_auditoria.MoveLast
            antpri = rs_auditoria!aud_ant
            clavepri = CLng(rs_auditoria!audnro)
            
            If claveult = clavepri Then
                seguir = True
            ElseIf actult <> antpri Then
                seguir = True
            End If
        End If
        
        rs_auditoria.Close
        
    End If

    If seguir Then
        tiene_evento = True
    Else
        tiene_evento = False
    End If
End Sub

Private Sub Indeve01(ByVal intnro As Integer, ByVal Empnro As Integer, ByVal ternro As Long, ByVal eventnro As Integer, ByVal topicnro As Integer, ByVal tfnro As Integer, ByRef v_tipovalor As Integer, ByRef v_valor As String, ByRef ok_valor As Boolean, ByRef msg_valor As String)
 Dim rs_event As New ADODB.Recordset

    v_tipovalor = 1
    
    StrSql = "SELECT * FROM event "
    StrSql = StrSql & " WHERE event.eventnro = " & eventnro
    OpenRecordset StrSql, rs_event
    If rs_event.EOF Then
        msg_valor = "No se ha encontrado el evento " & CStr(eventnro)
        ok_valor = False
    Else
        v_valor = rs_event!eventcode
        ok_valor = True
    End If
    
    rs_event.Close
    
End Sub

Private Sub Indeve02(ByVal intnro As Integer, ByVal Empnro As Integer, ByVal ternro As Long, ByVal eventnro As Integer, ByVal topicnro As Integer, ByVal tfnro As Integer, ByRef v_tipovalor As Integer, ByRef v_valor As String, ByRef ok_valor As Boolean, ByRef msg_valor As String)
 Dim rs_ter_doc As New ADODB.Recordset
 
    v_tipovalor = 1
    
    StrSql = "SELECT nrodoc FROM ter_doc "
    StrSql = StrSql & " INNER JOIN tipodocu ON ter_doc.tidnro = tipodocu.tidnro"
    StrSql = StrSql & " WHERE ter_doc.ternro = " & ternro
    StrSql = StrSql & " AND tipodocu.tidsigla = 'EMPLID'"
    OpenRecordset StrSql, rs_ter_doc
    If rs_ter_doc.EOF Then
        ok_valor = False
        msg_valor = "No se ha encontrado Registro de Tipo de Documento EMPLID para tercero " & CStr(ternro)
    Else
        ok_valor = True
        
        Call formato(CStr(rs_ter_doc!nrodoc), tfnro, v_tipovalor, v_valor)
        
    End If
    rs_ter_doc.Close
End Sub

Private Sub Indeve03(ByVal intnro As Integer, ByVal Empnro As Integer, ByVal ternro As Long, ByVal eventnro As Integer, ByVal topicnro As Integer, ByVal tfnro As Integer, ByRef v_tipovalor As Integer, ByRef v_valor As String, ByRef ok_valor As Boolean, ByRef msg_valor As String)
 Dim rs_intconfgen As New ADODB.Recordset
 Dim rs_topic_field As New ADODB.Recordset
 
    v_tipovalor = 1
    
    StrSql = "SELECT sysid FROM intconfgen "
    OpenRecordset StrSql, rs_intconfgen
    If rs_intconfgen.EOF Then
        ok_valor = False
        msg_valor = "No se ha encontrado Registro de Conf. General."
    Else
        ok_valor = True
        
        Call formato(CStr(rs_intconfgen!sysid), tfnro, v_tipovalor, v_valor)
        
    End If
    rs_intconfgen.Close
End Sub

Private Sub Indeve04(ByVal intnro As Integer, ByVal Empnro As Integer, ByVal ternro As Long, ByVal eventnro As Integer, ByVal topicnro As Integer, ByVal tfnro As Integer, ByRef v_tipovalor As Integer, ByRef v_valor As String, ByRef ok_valor As Boolean, ByRef msg_valor As String)
 Dim rs_empleado As New ADODB.Recordset
 Dim rs_topic_field As New ADODB.Recordset
 
    v_tipovalor = 1
    
    StrSql = "SELECT empleg FROM empleado WHERE ternro = " & ternro
    OpenRecordset StrSql, rs_empleado
    If rs_empleado.EOF Then
        ok_valor = False
        msg_valor = "No se ha encontrado Registro de Empleado."
    Else
        ok_valor = True
        
        Call formato(CStr(rs_empleado!empleg), tfnro, v_tipovalor, v_valor)
        
    End If
    
    rs_empleado.Close

End Sub

Private Sub Indeve05(ByVal intnro As Integer, ByVal Empnro As Integer, ByVal ternro As Long, ByVal eventnro As Integer, ByVal topicnro As Integer, ByVal tfnro As Integer, ByRef v_tipovalor As Integer, ByRef v_valor As String, ByRef ok_valor As Boolean, ByRef msg_valor As String)
 Dim rs_periodo As New ADODB.Recordset
 
    v_tipovalor = 1
    ok_valor = True
    msg_valor = ""
    
    ' Buscar periodo para la empresa actual
    StrSql = "SELECT * FROM periodo "
    StrSql = StrSql & " INNER JOIN interfaz ON periodo.pliqnro = interfaz.pliqnro WHERE interfaz.intnro = " & intnro
    OpenRecordset StrSql, rs_periodo
    If rs_periodo.EOF Then
        ok_valor = False
        msg_valor = "No se ha encontrado los datos del Período para la interface " & CStr(intnro)
    Else
        Call formato(CStr(rs_periodo!pliqhasta), tfnro, v_tipovalor, v_valor)
    End If
    
    rs_periodo.Close
    
End Sub

Private Sub Depeve01(ByVal intnro As Integer, ByVal Empnro As Integer, ByVal ternro As Long, ByVal eventnro As Integer, ByVal topicnro As Integer, ByVal tfnro As Integer, ByRef v_tipovalor As Integer, ByRef v_valor As String, ByRef ok_valor As Boolean, ByRef msg_valor As String)
 Dim rs_tercero As New ADODB.Recordset
 
    v_tipovalor = 1
    ok_valor = True
    msg_valor = ""
    
    ' DTC_XXX - [JOB] - BUSINESS_UNIT
    StrSql = "SELECT * FROM tercero WHERE ternro = " & ternro
    OpenRecordset StrSql, rs_tercero
    If rs_tercero.EOF Then
        ok_valor = False
        msg_valor = "No se ha encontrado el Registro de Tercero para ternro " & CStr(ternro)
    Else
        Call formato(CStr(rs_tercero!terape), tfnro, v_tipovalor, v_valor)
    End If
    
    rs_tercero.Close
    
End Sub

Private Sub Depeve02(ByVal intnro As Integer, ByVal Empnro As Integer, ByVal ternro As Long, ByVal eventnro As Integer, ByVal topicnro As Integer, ByVal tfnro As Integer, ByRef v_tipovalor As Integer, ByRef v_valor As String, ByRef ok_valor As Boolean, ByRef msg_valor As String)
 Dim rs_tercero As New ADODB.Recordset

    v_tipovalor = 1
    ok_valor = True
    msg_valor = ""
    
    ' DTC_XXX - [PERSO] - PATRONYMIC_NAME
    'StrSql = "SELECT * FROM tercero WHERE ternro = " & ternro
    'OpenRecordset StrSql, rs_tercero
    'If rs_tercero.EOF Then
    '    ok_valor = False
    '    msg_valor = "No se ha encontrado el Registro de Tercero para ternro " & CStr(ternro)
    'End If
    
    'If ok_valor Then
    '    If rs_tercero!tercasape = "" Then
    '        ok_valor = False
    '        msg_valor = "No tiene apellido de casada ternro " & CStr(ternro)
    '    End If
    'End If
    
    'If ok_valor Then
    '    Call formato(CStr(rs_tercero!tercasape), tfnro, v_tipovalor, v_valor)
    'End If
    
    'rs_tercero.Close
    
End Sub

Private Sub Depeve03(ByVal intnro As Integer, ByVal Empnro As Integer, ByVal ternro As Long, ByVal eventnro As Integer, ByVal topicnro As Integer, ByVal tfnro As Integer, ByRef v_tipovalor As Integer, ByRef v_valor As String, ByRef ok_valor As Boolean, ByRef msg_valor As String)
 Dim rs_tercero As New ADODB.Recordset
 
    v_tipovalor = 1
    ok_valor = True
    msg_valor = ""
    
    ' DTC_XXX - [PERSO] - FIRST_NAME
    StrSql = "SELECT * FROM tercero WHERE ternro = " & ternro
    OpenRecordset StrSql, rs_tercero
    If rs_tercero.EOF Then
        ok_valor = False
        msg_valor = "No se ha encontrado el Registro de Tercero para ternro " & CStr(ternro)
    Else
        Call formato(CStr(rs_tercero!ternom), tfnro, v_tipovalor, v_valor)
    End If
    
    rs_tercero.Close
    
End Sub

Private Sub Depeve04(ByVal intnro As Integer, ByVal Empnro As Integer, ByVal ternro As Long, ByVal eventnro As Integer, ByVal topicnro As Integer, ByVal tfnro As Integer, ByRef v_tipovalor As Integer, ByRef v_valor As String, ByRef ok_valor As Boolean, ByRef msg_valor As String)
 Dim Fecha As Date
 
 Dim rs_tercero As New ADODB.Recordset
 Dim rs_fases As New ADODB.Recordset
 
    v_tipovalor = 1
    ok_valor = True
    msg_valor = ""

    ' DTC_XXX - [PERSO] - HOME_COUNTRY
    StrSql = "SELECT empleado.empfaltagr FROM tercero "
    StrSql = StrSql & " INNER JOIN empleado ON tercero.ternro = empleado.ternro "
    StrSql = StrSql & " WHERE tercero.ternro = " & ternro
    OpenRecordset StrSql, rs_tercero
    If rs_tercero.EOF Then
        ok_valor = False
        msg_valor = "No se ha encontrado el Registro de Tercero para ternro " & CStr(ternro)
    Else
        StrSql = "SELECT altfec FROM fases "
        StrSql = StrSql & " WHERE fases.empleado = " & ternro
        StrSql = StrSql & " ORDER BY altfec DESC"
        OpenRecordset StrSql, rs_fases
        If Not rs_fases.EOF Then
            Fecha = rs_fases!altfec
        Else
            Fecha = rs_tercero!empfaltagr
        End If
        rs_fases.Close
        
        Call formato(CStr(Fecha), tfnro, v_tipovalor, v_valor)
        
    End If
    
    rs_tercero.Close
    
End Sub
Private Sub Depeve05(ByVal intnro As Integer, ByVal Empnro As Integer, ByVal ternro As Long, ByVal eventnro As Integer, ByVal topicnro As Integer, ByVal tfnro As Integer, ByRef v_tipovalor As Integer, ByRef v_valor As String, ByRef ok_valor As Boolean, ByRef msg_valor As String)
 Dim paisnro As Integer
 
 Dim rs_tercero As New ADODB.Recordset
 Dim rs_detdom As New ADODB.Recordset
 Dim rs_entity_value As New ADODB.Recordset
 
    v_tipovalor = 1
    ok_valor = True
    msg_valor = ""
    
    ' DTC_XXX - [PERSO] - HOME_COUNTRY
    StrSql = "SELECT * FROM tercero WHERE ternro = " & ternro
    OpenRecordset StrSql, rs_tercero
    If rs_tercero.EOF Then
        ok_valor = False
        msg_valor = "No se ha encontrado el Registro de Tercero para ternro " & CStr(ternro)
    End If
    rs_tercero.Close
    
    If ok_valor Then
        StrSql = "SELECT detdom.paisnro FROM cabdom"
        StrSql = StrSql & " INNER JOIN detdom ON cabdom.domnro = detdom.domnro "
        StrSql = StrSql & " WHERE cabdom.ternro = " & ternro & " AND cabdom.tidonro = 2"
        OpenRecordset StrSql, rs_detdom
        If rs_detdom.EOF Then
            ok_valor = False
            msg_valor = "No tiene domicilio Particular. ternro " & CStr(ternro)
        Else
            paisnro = rs_detdom!paisnro
        End If
        rs_detdom.Close
    End If
    
    If ok_valor Then
        StrSql = "SELECT entity_value.codetosend FROM entity_map_value "
        StrSql = StrSql & " INNER JOIN topic_field ON topic_field.entmapnro = entity_map_value.entmapnro "
        StrSql = StrSql & " INNER JOIN entity_value ON entity_value.entvalnro = entity_map_value.imgvalor "
        StrSql = StrSql & " WHERE entity_map_value.domvalor = " & paisnro
        StrSql = StrSql & " AND topic_field.tfnro = " & tfnro
        'StrSql = StrSql & " AND entity_map_value.empnro = " & empnro
        OpenRecordset StrSql, rs_entity_value
        If rs_entity_value.EOF Then
            ok_valor = False
        Else
            v_valor = rs_entity_value!codetosend
        End If
        rs_entity_value.Close
    End If

End Sub

Private Sub Depeve06(ByVal intnro As Integer, ByVal Empnro As Integer, ByVal ternro As Long, ByVal eventnro As Integer, ByVal topicnro As Integer, ByVal tfnro As Integer, ByRef v_tipovalor As Integer, ByRef v_valor As String, ByRef ok_valor As Boolean, ByRef msg_valor As String)
 Dim calle As String
 
 Dim rs_tercero As New ADODB.Recordset
 Dim rs_detdom As New ADODB.Recordset
 Dim rs_entity_value As New ADODB.Recordset
 
    v_tipovalor = 1
    ok_valor = True
    msg_valor = ""
    
    ' DTC_XXX - [PERSO] - HOME_ADDRESS1
    StrSql = "SELECT * FROM tercero WHERE ternro = " & ternro
    OpenRecordset StrSql, rs_tercero
    If rs_tercero.EOF Then
        ok_valor = False
        msg_valor = "No se ha encontrado el Registro de Tercero para ternro " & CStr(ternro)
    End If
    rs_tercero.Close
    
    If ok_valor Then
        StrSql = "SELECT detdom.calle FROM cabdom"
        StrSql = StrSql & " INNER JOIN detdom ON cabdom.domnro = detdom.domnro "
        StrSql = StrSql & " WHERE cabdom.ternro = " & ternro & " AND cabdom.tidonro = 2"
        OpenRecordset StrSql, rs_detdom
        If rs_detdom.EOF Then
            ok_valor = False
            msg_valor = "No tiene domicilio Particular. ternro " & CStr(ternro)
        Else
            calle = rs_detdom!calle
            Call formato(calle, tfnro, v_tipovalor, v_valor)
        End If
        
        rs_detdom.Close
    End If
End Sub

Private Sub Depeve07(ByVal intnro As Integer, ByVal Empnro As Integer, ByVal ternro As Long, ByVal eventnro As Integer, ByVal topicnro As Integer, ByVal tfnro As Integer, ByRef v_tipovalor As Integer, ByRef v_valor As String, ByRef ok_valor As Boolean, ByRef msg_valor As String)
 Dim nro As String
 
 Dim rs_tercero As New ADODB.Recordset
 Dim rs_detdom As New ADODB.Recordset
 Dim rs_entity_value As New ADODB.Recordset
 
    v_tipovalor = 1
    ok_valor = True
    msg_valor = ""
    
    ' DTC_XXX - [PERSO] - HOME_ADDRESS2
    StrSql = "SELECT * FROM tercero WHERE ternro = " & ternro
    OpenRecordset StrSql, rs_tercero
    If rs_tercero.EOF Then
        ok_valor = False
        msg_valor = "No se ha encontrado el Registro de Tercero para ternro " & CStr(ternro)
    End If
    rs_tercero.Close
    
    If ok_valor Then
        StrSql = "SELECT detdom.nro FROM cabdom"
        StrSql = StrSql & " INNER JOIN detdom ON cabdom.domnro = detdom.domnro "
        StrSql = StrSql & " WHERE cabdom.ternro = " & ternro & " AND cabdom.tidonro = 2"
        OpenRecordset StrSql, rs_detdom
        If rs_detdom.EOF Then
            ok_valor = False
            msg_valor = "No tiene domicilio Particular. ternro " & CStr(ternro)
        Else
            nro = rs_detdom!nro
            Call formato(nro, tfnro, v_tipovalor, v_valor)
        End If
        
        rs_detdom.Close
    End If
End Sub

Private Sub Depeve08(ByVal intnro As Integer, ByVal Empnro As Integer, ByVal ternro As Long, ByVal eventnro As Integer, ByVal topicnro As Integer, ByVal tfnro As Integer, ByRef v_tipovalor As Integer, ByRef v_valor As String, ByRef ok_valor As Boolean, ByRef msg_valor As String)
 Dim piso As String
 
 Dim rs_tercero As New ADODB.Recordset
 Dim rs_detdom As New ADODB.Recordset
 Dim rs_entity_value As New ADODB.Recordset
 
    v_tipovalor = 1
    ok_valor = True
    msg_valor = ""
    
    ' DTC_XXX - [PERSO] - HOME_ADDRESS3
    StrSql = "SELECT * FROM tercero WHERE ternro = " & ternro
    OpenRecordset StrSql, rs_tercero
    If rs_tercero.EOF Then
        ok_valor = False
        msg_valor = "No se ha encontrado el Registro de Tercero para ternro " & CStr(ternro)
    End If
    rs_tercero.Close
    
    If ok_valor Then
        StrSql = "SELECT detdom.piso FROM cabdom"
        StrSql = StrSql & " INNER JOIN detdom ON cabdom.domnro = detdom.domnro "
        StrSql = StrSql & " WHERE cabdom.ternro = " & ternro & " AND cabdom.tidonro = 2"
        OpenRecordset StrSql, rs_detdom
        If rs_detdom.EOF Then
            ok_valor = False
            msg_valor = "No tiene domicilio Particular. ternro " & CStr(ternro)
        Else
            piso = rs_detdom!piso
            Call formato(piso, tfnro, v_tipovalor, v_valor)
        End If
        
        rs_detdom.Close
    End If
End Sub

Private Sub Depeve09(ByVal intnro As Integer, ByVal Empnro As Integer, ByVal ternro As Long, ByVal eventnro As Integer, ByVal topicnro As Integer, ByVal tfnro As Integer, ByRef v_tipovalor As Integer, ByRef v_valor As String, ByRef ok_valor As Boolean, ByRef msg_valor As String)
 Dim oficdepto As String
 
 Dim rs_tercero As New ADODB.Recordset
 Dim rs_detdom As New ADODB.Recordset
 Dim rs_entity_value As New ADODB.Recordset
 
    v_tipovalor = 1
    ok_valor = True
    msg_valor = ""
    
    ' DTC_XXX - [PERSO] - HOME_ADDRESS4
    StrSql = "SELECT * FROM tercero WHERE ternro = " & ternro
    OpenRecordset StrSql, rs_tercero
    If rs_tercero.EOF Then
        ok_valor = False
        msg_valor = "No se ha encontrado el Registro de Tercero para ternro " & CStr(ternro)
    End If
    rs_tercero.Close
    
    If ok_valor Then
        StrSql = "SELECT detdom.oficdepto FROM cabdom"
        StrSql = StrSql & " INNER JOIN detdom ON cabdom.domnro = detdom.domnro "
        StrSql = StrSql & " WHERE cabdom.ternro = " & ternro & " AND cabdom.tidonro = 2"
        OpenRecordset StrSql, rs_detdom
        If rs_detdom.EOF Then
            ok_valor = False
            msg_valor = "No tiene domicilio Particular. ternro " & CStr(ternro)
        Else
            oficdepto = rs_detdom!oficdepto
            Call formato(oficdepto, tfnro, v_tipovalor, v_valor)
        End If
        
        rs_detdom.Close
    End If
End Sub

Private Sub Depeve10(ByVal intnro As Integer, ByVal Empnro As Integer, ByVal ternro As Long, ByVal eventnro As Integer, ByVal topicnro As Integer, ByVal tfnro As Integer, ByRef v_tipovalor As Integer, ByRef v_valor As String, ByRef ok_valor As Boolean, ByRef msg_valor As String)
 Dim locdesc As String
 
 Dim rs_tercero As New ADODB.Recordset
 Dim rs_detdom As New ADODB.Recordset
 Dim rs_entity_value As New ADODB.Recordset
 
    v_tipovalor = 1
    ok_valor = True
    msg_valor = ""
    
    ' DTC_XXX - [PERSO] - HOME_CITY
    StrSql = "SELECT * FROM tercero WHERE ternro = " & ternro
    OpenRecordset StrSql, rs_tercero
    If rs_tercero.EOF Then
        ok_valor = False
        msg_valor = "No se ha encontrado el Registro de Tercero para ternro " & CStr(ternro)
    End If
    rs_tercero.Close

    If ok_valor Then
        StrSql = "SELECT localidad.locdesc FROM cabdom"
        StrSql = StrSql & " INNER JOIN detdom ON cabdom.domnro = detdom.domnro "
        StrSql = StrSql & " INNER JOIN localidad ON localidad.locnro = detdom.locnro "
        StrSql = StrSql & " WHERE cabdom.ternro = " & ternro & " AND cabdom.tidonro = 2"
        OpenRecordset StrSql, rs_detdom
        If rs_detdom.EOF Then
            ok_valor = False
            msg_valor = "No tiene Localidad asociada. ternro " & CStr(ternro)
        Else
            locdesc = rs_detdom!locdesc
            Call formato(locdesc, tfnro, v_tipovalor, v_valor)
        End If
        
        rs_detdom.Close
    End If
End Sub

Private Sub Depeve11(ByVal intnro As Integer, ByVal Empnro As Integer, ByVal ternro As Long, ByVal eventnro As Integer, ByVal topicnro As Integer, ByVal tfnro As Integer, ByRef v_tipovalor As Integer, ByRef v_valor As String, ByRef ok_valor As Boolean, ByRef msg_valor As String)
 Dim codigopostal As String
 
 Dim rs_tercero As New ADODB.Recordset
 Dim rs_detdom As New ADODB.Recordset
 Dim rs_entity_value As New ADODB.Recordset
 
    v_tipovalor = 1
    ok_valor = True
    msg_valor = ""
    
    ' DTC_XXX - [PERSO] - HOME_CITY
    StrSql = "SELECT * FROM tercero WHERE ternro = " & ternro
    OpenRecordset StrSql, rs_tercero
    If rs_tercero.EOF Then
        ok_valor = False
        msg_valor = "No se ha encontrado el Registro de Tercero para ternro " & CStr(ternro)
    End If
    rs_tercero.Close
    
    If ok_valor Then
        StrSql = "SELECT detdom.codigopostal FROM cabdom"
        StrSql = StrSql & " INNER JOIN detdom ON cabdom.domnro = detdom.domnro "
        StrSql = StrSql & " WHERE cabdom.ternro = " & ternro & " AND cabdom.tidonro = 2"
        OpenRecordset StrSql, rs_detdom
        If rs_detdom.EOF Then
            ok_valor = False
            msg_valor = "No tiene domicilio Particular. ternro " & CStr(ternro)
            codigopostal = ""
        Else
            codigopostal = rs_detdom!codigopostal
        End If
        
        If EsNulo(codigopostal) Or codigopostal = "0" Then
            ok_valor = False
            msg_valor = "No tiene CP asociado. ternro " & CStr(ternro)
        Else
            Call formato(codigopostal, tfnro, v_tipovalor, v_valor)
        End If
        
        rs_detdom.Close
    End If
End Sub

Private Sub Depeve12(ByVal intnro As Integer, ByVal Empnro As Integer, ByVal ternro As Long, ByVal eventnro As Integer, ByVal topicnro As Integer, ByVal tfnro As Integer, ByRef v_tipovalor As Integer, ByRef v_valor As String, ByRef ok_valor As Boolean, ByRef msg_valor As String)
 Dim provdesc As String
 
 Dim rs_tercero As New ADODB.Recordset
 Dim rs_detdom As New ADODB.Recordset
 Dim rs_entity_value As New ADODB.Recordset
 
    v_tipovalor = 1
    ok_valor = True
    msg_valor = ""
    
    ' DTC_XXX - [PERSO] - HOME_CITY
    StrSql = "SELECT * FROM tercero WHERE ternro = " & ternro
    OpenRecordset StrSql, rs_tercero
    If rs_tercero.EOF Then
        ok_valor = False
        msg_valor = "No se ha encontrado el Registro de Tercero para ternro " & CStr(ternro)
    End If
    rs_tercero.Close

    If ok_valor Then
        StrSql = "SELECT provincia.provdesc FROM cabdom"
        StrSql = StrSql & " INNER JOIN detdom ON cabdom.domnro = detdom.domnro "
        StrSql = StrSql & " INNER JOIN provincia ON provincia.provnro = detdom.provnro "
        StrSql = StrSql & " WHERE cabdom.ternro = " & ternro & " AND cabdom.tidonro = 2"
        OpenRecordset StrSql, rs_detdom
        If rs_detdom.EOF Then
            ok_valor = False
            msg_valor = "No tiene Provincia asociada. ternro " & CStr(ternro)
        Else
            provdesc = rs_detdom!provdesc
            Call formato(provdesc, tfnro, v_tipovalor, v_valor)
        End If
        
        rs_detdom.Close
    End If
End Sub

Private Sub Depeve21(ByVal intnro As Integer, ByVal Empnro As Integer, ByVal ternro As Long, ByVal eventnro As Integer, ByVal topicnro As Integer, ByVal tfnro As Integer, ByRef v_tipovalor As Integer, ByRef v_valor As String, ByRef ok_valor As Boolean, ByRef msg_valor As String)
 Dim calle As String
 
 Dim rs_tercero As New ADODB.Recordset
 Dim rs_detdom As New ADODB.Recordset
 Dim rs_entity_value As New ADODB.Recordset
 
    v_tipovalor = 1
    ok_valor = True
    msg_valor = ""
    
    ' DTC_XXX - [PERSO] - SEX
    StrSql = "SELECT * FROM tercero WHERE ternro = " & ternro
    OpenRecordset StrSql, rs_tercero
    If rs_tercero.EOF Then
        ok_valor = False
        msg_valor = "No se ha encontrado el Registro de Tercero para ternro " & CStr(ternro)
    Else
        If CInt(rs_tercero!tersex) = -1 Then
            v_valor = "M"
        Else
            v_valor = "F"
        End If
    End If
    
    rs_tercero.Close
    
End Sub

Private Sub Depeve22(ByVal intnro As Integer, ByVal Empnro As Integer, ByVal ternro As Long, ByVal eventnro As Integer, ByVal topicnro As Integer, ByVal tfnro As Integer, ByRef v_tipovalor As Integer, ByRef v_valor As String, ByRef ok_valor As Boolean, ByRef msg_valor As String)
 
 Dim rs_tercero As New ADODB.Recordset
 Dim rs_detdom As New ADODB.Recordset
 Dim rs_entity_value As New ADODB.Recordset
 
    v_tipovalor = 1
    ok_valor = True
    msg_valor = ""
    
    ' DTC_XXX - [PERSO] - MAR_STATUS
    StrSql = "SELECT * FROM tercero WHERE ternro = " & ternro
    OpenRecordset StrSql, rs_tercero
    If rs_tercero.EOF Then
        ok_valor = False
        msg_valor = "No se ha encontrado el Registro de Tercero para ternro " & CStr(ternro)
    Else
        StrSql = "SELECT entity_value.codetosend FROM entity_map_value "
        StrSql = StrSql & " INNER JOIN topic_field ON topic_field.entmapnro = entity_map_value.entmapnro "
        StrSql = StrSql & " INNER JOIN entity_value ON entity_value.entvalnro = entity_map_value.imgvalor "
        StrSql = StrSql & " WHERE entity_map_value.domvalor = " & rs_tercero!estcivnro
        StrSql = StrSql & " AND topic_field.tfnro = " & tfnro
        'StrSql = StrSql & " AND entity_map_value.empnro = " & empnro
        OpenRecordset StrSql, rs_entity_value
        If rs_entity_value.EOF Then
            ok_valor = False
        Else
            v_valor = rs_entity_value!codetosend
        End If
        rs_entity_value.Close
    End If
        
    rs_tercero.Close
    
End Sub

Private Sub Depeve23(ByVal intnro As Integer, ByVal Empnro As Integer, ByVal ternro As Long, ByVal eventnro As Integer, ByVal topicnro As Integer, ByVal tfnro As Integer, ByRef v_tipovalor As Integer, ByRef v_valor As String, ByRef ok_valor As Boolean, ByRef msg_valor As String)
 Dim rs_tercero As New ADODB.Recordset

    v_tipovalor = 1
    ok_valor = True
    msg_valor = ""
    
    ' DTC_XXX - [PERSO] - BIRTHDATE
    StrSql = "SELECT terfecnac FROM tercero WHERE ternro = " & ternro
    OpenRecordset StrSql, rs_tercero
    If rs_tercero.EOF Then
        ok_valor = False
        msg_valor = "No se ha encontrado el Registro de Tercero para ternro " & CStr(ternro)
    Else
        Call formato(CStr(rs_tercero!terfecnac), tfnro, v_tipovalor, v_valor)
    End If
    
    rs_tercero.Close
    
End Sub

Private Sub Depeve24(ByVal intnro As Integer, ByVal Empnro As Integer, ByVal ternro As Long, ByVal eventnro As Integer, ByVal topicnro As Integer, ByVal tfnro As Integer, ByRef v_tipovalor As Integer, ByRef v_valor As String, ByRef ok_valor As Boolean, ByRef msg_valor As String)
 Dim nivnro As Integer
 
 Dim rs_tercero As New ADODB.Recordset
 Dim rs_detdom As New ADODB.Recordset
 Dim rs_entity_value As New ADODB.Recordset
 Dim rs_cap_estformal As New ADODB.Recordset
 
    v_tipovalor = 1
    ok_valor = True
    msg_valor = ""
    
    ' DTC_XXX - [PERSO] - HIGHEST_EDUC_LVL
    StrSql = "SELECT * FROM tercero WHERE ternro = " & ternro
    OpenRecordset StrSql, rs_tercero
    If rs_tercero.EOF Then
        ok_valor = False
        msg_valor = "No se ha encontrado el Registro de Tercero para ternro " & CStr(ternro)
    End If
    rs_tercero.Close
    
    If ok_valor Then
        StrSql = "SELECT nivnro FROM cap_estformal "
        StrSql = StrSql & " WHERE ternro = " & ternro & " AND capactual = -1 "
        OpenRecordset StrSql, rs_cap_estformal
        If rs_cap_estformal.EOF Then
            ok_valor = False
            msg_valor = "No se ha encontrado el nivel de estudio (capactual=-1) para ternro " & CStr(ternro)
        Else
            nivnro = rs_cap_estformal!nivnro
        End If
        rs_cap_estformal.Close
    End If
    
    If ok_valor Then
        StrSql = "SELECT entity_value.codetosend FROM entity_map_value "
        StrSql = StrSql & " INNER JOIN topic_field ON topic_field.entmapnro = entity_map_value.entmapnro "
        StrSql = StrSql & " INNER JOIN entity_value ON entity_value.entvalnro = entity_map_value.imgvalor "
        StrSql = StrSql & " WHERE entity_map_value.domvalor = " & nivnro
        StrSql = StrSql & " AND topic_field.tfnro = " & tfnro
        'StrSql = StrSql & " AND entity_map_value.empnro = " & empnro
        OpenRecordset StrSql, rs_entity_value
        If rs_entity_value.EOF Then
            ok_valor = False
        Else
            v_valor = rs_entity_value!codetosend
        End If
        rs_entity_value.Close
    End If
        
End Sub

Private Sub Depeve25(ByVal intnro As Integer, ByVal Empnro As Integer, ByVal ternro As Long, ByVal eventnro As Integer, ByVal topicnro As Integer, ByVal tfnro As Integer, ByRef v_tipovalor As Integer, ByRef v_valor As String, ByRef ok_valor As Boolean, ByRef msg_valor As String)
 
 Dim rs_tercero As New ADODB.Recordset
 Dim rs_detdom As New ADODB.Recordset
 Dim rs_entity_value As New ADODB.Recordset
 
    v_tipovalor = 1
    ok_valor = True
    msg_valor = ""
    
    ' DTC_XXX - [PERSO] - CITIZENSHIP_1
    StrSql = "SELECT paisnro FROM tercero WHERE ternro = " & ternro
    OpenRecordset StrSql, rs_tercero
    If rs_tercero.EOF Then
        ok_valor = False
        msg_valor = "No se ha encontrado el Registro de Tercero para ternro " & CStr(ternro)
    Else
        StrSql = "SELECT entity_value.codetosend FROM entity_map_value "
        StrSql = StrSql & " INNER JOIN topic_field ON topic_field.entmapnro = entity_map_value.entmapnro "
        StrSql = StrSql & " INNER JOIN entity_value ON entity_value.entvalnro = entity_map_value.imgvalor "
        StrSql = StrSql & " WHERE entity_map_value.domvalor = " & rs_tercero!paisnro
        StrSql = StrSql & " AND topic_field.tfnro = " & tfnro
        'StrSql = StrSql & " AND entity_map_value.empnro = " & empnro
        OpenRecordset StrSql, rs_entity_value
        If rs_entity_value.EOF Then
            ok_valor = False
        Else
            v_valor = rs_entity_value!codetosend
        End If
        rs_entity_value.Close
    End If
        
    rs_tercero.Close
    
End Sub

Private Sub Depeve26(ByVal intnro As Integer, ByVal Empnro As Integer, ByVal ternro As Long, ByVal eventnro As Integer, ByVal topicnro As Integer, ByVal tfnro As Integer, ByRef v_tipovalor As Integer, ByRef v_valor As String, ByRef ok_valor As Boolean, ByRef msg_valor As String)
 Dim entnro As Integer
 
 Dim rs_tercero As New ADODB.Recordset
 Dim rs_entity As New ADODB.Recordset
 Dim rs_emp_attrib_value As New ADODB.Recordset
 
    v_tipovalor = 1
    ok_valor = True
    msg_valor = ""
    
    ' DTC_XXX - [JOB] - BUSINESS_UNIT
    StrSql = "SELECT * FROM tercero WHERE ternro = " & ternro
    OpenRecordset StrSql, rs_tercero
    If rs_tercero.EOF Then
        ok_valor = False
        msg_valor = "No se ha encontrado el Registro de Tercero para ternro " & CStr(ternro)
    End If
    rs_tercero.Close
    
    If ok_valor Then
        ' Buscar valor configurado a mano
        StrSql = "SELECT * FROM entity "
        StrSql = StrSql & " INNER JOIN topic_field ON entity.entnro = topic_field.entnro "
        StrSql = StrSql & " WHERE topic_field.tfnro = " & tfnro
        OpenRecordset StrSql, rs_entity
        If rs_entity.EOF Then
            ok_valor = False
            msg_valor = "No esta conf. Entidad asociada."
        Else
            entnro = rs_entity!entnro
        End If
        rs_entity.Close
    End If
    
    If ok_valor Then
        ' Buscar valor asociado
        StrSql = "SELECT entity_value.codetosend, emp_attrib_value.entvalnro FROM emp_attrib_value "
        StrSql = StrSql & " INNER JOIN entity_value ON emp_attrib_value.entvalnro = entity_value.entvalnro "
        StrSql = StrSql & " WHERE emp_attrib_value.ternro = " & ternro
        'StrSql = StrSql & " AND emp_attrib_value.empnro = " & empnro
        StrSql = StrSql & " AND entity_value.entnro = " & entnro
        OpenRecordset StrSql, rs_emp_attrib_value
        If rs_emp_attrib_value.EOF Then
            ok_valor = False
            msg_valor = "No esta conf. Manual para campo " & CStr(tfnro) & ". ternro " & CStr(ternro)
        Else
            v_valor = rs_emp_attrib_value!codetosend
        End If
        rs_emp_attrib_value.Close
    End If
    
End Sub

Private Sub Depeve27(ByVal intnro As Integer, ByVal Empnro As Integer, ByVal ternro As Long, ByVal eventnro As Integer, ByVal topicnro As Integer, ByVal tfnro As Integer, ByRef v_tipovalor As Integer, ByRef v_valor As String, ByRef ok_valor As Boolean, ByRef msg_valor As String)
 Dim ternro_estr As Integer
 Dim intfhasta As Date
 Dim intfdesde As Date
 
 Dim rs_tercero As New ADODB.Recordset
 Dim rs_interfaz As New ADODB.Recordset
 Dim rs_his_estructura As New ADODB.Recordset
 Dim rs_entity_value As New ADODB.Recordset
 
    v_tipovalor = 1
    ok_valor = True
    msg_valor = ""

    ' DTC_XXX - [JOB] - LOCATION
    StrSql = "SELECT * FROM tercero WHERE ternro = " & ternro
    OpenRecordset StrSql, rs_tercero
    If rs_tercero.EOF Then
        ok_valor = False
        msg_valor = "No se ha encontrado el Registro de Tercero para ternro " & CStr(ternro)
    End If
    rs_tercero.Close
    
    If ok_valor Then
        StrSql = "SELECT * FROM interfaz WHERE intnro = " & intnro
        OpenRecordset StrSql, rs_interfaz
        If rs_interfaz.EOF Then
            ok_valor = False
        Else
            intfhasta = rs_interfaz!intfhasta
            intfdesde = rs_interfaz!intfdesde
        End If
        rs_interfaz.Close
    End If
    
    If ok_valor Then
        StrSql = "SELECT sucursal.ternro FROM his_estructura "
        StrSql = StrSql & " INNER JOIN sucursal ON his_estructura.estrnro = sucursal.estrnro "
        StrSql = StrSql & " WHERE his_estructura.ternro = " & ternro
        StrSql = StrSql & " AND his_estructura.tenro = 1"
        StrSql = StrSql & " AND his_estructura.htetdesde <= " & ConvFecha(intfhasta)
        StrSql = StrSql & " AND (his_estructura.htethasta >= " & ConvFecha(intfdesde)
        StrSql = StrSql & " OR his_estructura.htethasta is null)"
        StrSql = StrSql & " ORDER BY his_estructura.htetdesde DESC"
        OpenRecordset StrSql, rs_his_estructura
        If rs_his_estructura.EOF Then
            ok_valor = False
            msg_valor = "No tiene sucursal(1). ternro " & CStr(ternro)
        Else
            ternro_estr = rs_his_estructura!ternro
        End If
        rs_his_estructura.Close
    End If
    
    If ok_valor Then
        StrSql = "SELECT entity_value.codetosend FROM entity_map_value "
        StrSql = StrSql & " INNER JOIN topic_field ON topic_field.entmapnro = entity_map_value.entmapnro "
        StrSql = StrSql & " INNER JOIN entity_value ON entity_value.entvalnro = entity_map_value.imgvalor "
        StrSql = StrSql & " WHERE entity_map_value.domvalor = " & ternro_estr
        StrSql = StrSql & " AND topic_field.tfnro = " & tfnro
        'StrSql = StrSql & " AND entity_map_value.empnro = " & empnro
        OpenRecordset StrSql, rs_entity_value
        If rs_entity_value.EOF Then
            ok_valor = False
        Else
            v_valor = rs_entity_value!codetosend
        End If
        rs_entity_value.Close
    End If
        
End Sub

Private Sub Depeve28(ByVal intnro As Integer, ByVal Empnro As Integer, ByVal ternro As Long, ByVal eventnro As Integer, ByVal topicnro As Integer, ByVal tfnro As Integer, ByRef v_tipovalor As Integer, ByRef v_valor As String, ByRef ok_valor As Boolean, ByRef msg_valor As String)
 Dim estrnro As Integer
 Dim intfhasta As Date
 Dim intfdesde As Date
 
 Dim rs_tercero As New ADODB.Recordset
 Dim rs_interfaz As New ADODB.Recordset
 Dim rs_his_estructura As New ADODB.Recordset
 Dim rs_entity_value As New ADODB.Recordset
 
    v_tipovalor = 1
    ok_valor = True
    msg_valor = ""

    ' DTC_XXX - [JOB] - CONTRACT_TYPE
    StrSql = "SELECT * FROM tercero WHERE ternro = " & ternro
    OpenRecordset StrSql, rs_tercero
    If rs_tercero.EOF Then
        ok_valor = False
        msg_valor = "No se ha encontrado el Registro de Tercero para ternro " & CStr(ternro)
    End If
    rs_tercero.Close
    
    If ok_valor Then
        StrSql = "SELECT * FROM interfaz WHERE intnro = " & intnro
        OpenRecordset StrSql, rs_interfaz
        If rs_interfaz.EOF Then
            ok_valor = False
        Else
            intfhasta = rs_interfaz!intfhasta
            intfdesde = rs_interfaz!intfdesde
        End If
        rs_interfaz.Close
    End If
    
    If ok_valor Then
        StrSql = "SELECT tipocont.tcnro FROM his_estructura "
        StrSql = StrSql & " INNER JOIN tipocont ON his_estructura.estrnro = tipocont.estrnro "
        StrSql = StrSql & " WHERE his_estructura.ternro = " & ternro
        StrSql = StrSql & " AND his_estructura.tenro = 18"
        StrSql = StrSql & " AND his_estructura.htetdesde <= " & ConvFecha(intfhasta)
        StrSql = StrSql & " AND (his_estructura.htethasta >= " & ConvFecha(intfdesde)
        StrSql = StrSql & " OR his_estructura.htethasta is null)"
        StrSql = StrSql & " ORDER BY his_estructura.htetdesde DESC"
        OpenRecordset StrSql, rs_his_estructura
        If rs_his_estructura.EOF Then
            ok_valor = False
            msg_valor = "No tiene Contrato Actual(18). ternro " & CStr(ternro)
        Else
            estrnro = rs_his_estructura!tcnro
        End If
        rs_his_estructura.Close
    End If
    
    If ok_valor Then
        StrSql = "SELECT entity_value.codetosend FROM entity_map_value "
        StrSql = StrSql & " INNER JOIN topic_field ON topic_field.entmapnro = entity_map_value.entmapnro "
        StrSql = StrSql & " INNER JOIN entity_value ON entity_value.entvalnro = entity_map_value.imgvalor "
        StrSql = StrSql & " WHERE entity_map_value.domvalor = " & estrnro
        StrSql = StrSql & " AND topic_field.tfnro = " & tfnro
        'StrSql = StrSql & " AND entity_map_value.empnro = " & empnro
        OpenRecordset StrSql, rs_entity_value
        If rs_entity_value.EOF Then
            ok_valor = False
        Else
            v_valor = rs_entity_value!codetosend
        End If
        rs_entity_value.Close
    End If
        
End Sub

Private Sub Depeve29(ByVal intnro As Integer, ByVal Empnro As Integer, ByVal ternro As Long, ByVal eventnro As Integer, ByVal topicnro As Integer, ByVal tfnro As Integer, ByRef v_tipovalor As Integer, ByRef v_valor As String, ByRef ok_valor As Boolean, ByRef msg_valor As String)
 Dim estrnro As Integer
 Dim intfhasta As Date
 Dim intfdesde As Date
 
 Dim rs_tercero As New ADODB.Recordset
 Dim rs_interfaz As New ADODB.Recordset
 Dim rs_his_estructura As New ADODB.Recordset
 Dim rs_his_estructura_ant As New ADODB.Recordset
 Dim rs_entity_value As New ADODB.Recordset
 
    v_tipovalor = 1
    ok_valor = True
    msg_valor = ""

    ' DTC_XXX - [JOB] - CONTRACT_TYPE
    StrSql = "SELECT * FROM tercero WHERE ternro = " & ternro
    OpenRecordset StrSql, rs_tercero
    If rs_tercero.EOF Then
        ok_valor = False
        msg_valor = "No se ha encontrado el Registro de Tercero para ternro " & CStr(ternro)
    End If
    rs_tercero.Close
    
    If ok_valor Then
        StrSql = "SELECT * FROM interfaz WHERE intnro = " & intnro
        OpenRecordset StrSql, rs_interfaz
        If rs_interfaz.EOF Then
            ok_valor = False
        Else
            intfhasta = rs_interfaz!intfhasta
            intfdesde = rs_interfaz!intfdesde
        End If
        rs_interfaz.Close
    End If
    
    If ok_valor Then
        ' Busco la estructura actual (contrato actual)
        StrSql = "SELECT htetdesde, estrnro FROM his_estructura "
        StrSql = StrSql & " WHERE ternro = " & ternro
        StrSql = StrSql & " AND tenro = 18"
        StrSql = StrSql & " AND htetdesde <= " & ConvFecha(intfhasta)
        StrSql = StrSql & " AND (htethasta >= " & ConvFecha(intfdesde) & " OR htethasta is null)"
        StrSql = StrSql & " ORDER BY htetdesde DESC"
        OpenRecordset StrSql, rs_his_estructura
        If rs_his_estructura.EOF Then
            ok_valor = False
            msg_valor = "No tiene Contrato Actual(18). ternro " & CStr(ternro)
        Else
            ' Busco la estructura anterior (contrato anterior)
            StrSql = "SELECT htethasta FROM his_estructura "
            StrSql = StrSql & " WHERE ternro = " & ternro
            StrSql = StrSql & " AND tenro = 18"
            StrSql = StrSql & " AND htetdesde <= " & ConvFecha(rs_his_estructura!htetdesde)
            StrSql = StrSql & " AND estrnro <> " & rs_his_estructura!estrnro
            StrSql = StrSql & " ORDER BY htetdesde DESC"
            OpenRecordset StrSql, rs_his_estructura_ant
            If Not rs_his_estructura_ant.EOF Then
                Call formato(CStr(rs_his_estructura_ant!htethasta), tfnro, v_tipovalor, v_valor)
            Else
                Call Depeve04(intnro, Empnro, ternro, eventnro, topicnro, tfnro, v_tipovalor, v_valor, ok_valor, msg_valor)
            End If

            rs_his_estructura_ant.Close

        End If
        rs_his_estructura.Close
    End If
    
End Sub

Private Sub Depeve30(ByVal intnro As Integer, ByVal Empnro As Integer, ByVal ternro As Long, ByVal eventnro As Integer, ByVal topicnro As Integer, ByVal tfnro As Integer, ByRef v_tipovalor As Integer, ByRef v_valor As String, ByRef ok_valor As Boolean, ByRef msg_valor As String)
 Dim htethasta As String
 Dim intfhasta As Date
 Dim intfdesde As Date
 
 Dim rs_tercero As New ADODB.Recordset
 Dim rs_interfaz As New ADODB.Recordset
 Dim rs_his_estructura As New ADODB.Recordset
 Dim rs_entity_value As New ADODB.Recordset
 
    v_tipovalor = 1
    ok_valor = True
    msg_valor = ""

    ' DTC_XXX - [JOB] - CONTRACT_TYPE
    StrSql = "SELECT * FROM tercero WHERE ternro = " & ternro
    OpenRecordset StrSql, rs_tercero
    If rs_tercero.EOF Then
        ok_valor = False
        msg_valor = "No se ha encontrado el Registro de Tercero para ternro " & CStr(ternro)
    End If
    rs_tercero.Close
    
    If ok_valor Then
        StrSql = "SELECT * FROM interfaz WHERE intnro = " & intnro
        OpenRecordset StrSql, rs_interfaz
        If rs_interfaz.EOF Then
            ok_valor = False
        Else
            intfhasta = rs_interfaz!intfhasta
            intfdesde = rs_interfaz!intfdesde
        End If
        rs_interfaz.Close
    End If
    
    If ok_valor Then
        StrSql = "SELECT htethasta FROM his_estructura "
        StrSql = StrSql & " WHERE ternro = " & ternro
        StrSql = StrSql & " AND tenro = 18"
        StrSql = StrSql & " AND htetdesde <= " & ConvFecha(intfhasta)
        StrSql = StrSql & " AND (htethasta >= " & ConvFecha(intfdesde) & " OR htethasta is null)"
        StrSql = StrSql & " ORDER BY htetdesde DESC"
        OpenRecordset StrSql, rs_his_estructura
        If rs_his_estructura.EOF Then
            ok_valor = False
            msg_valor = "No tiene Contrato Actual(18). ternro " & CStr(ternro)
        Else
            If EsNulo(rs_his_estructura!htethasta) Then
                htethasta = ""
            Else
                htethasta = rs_his_estructura!htethasta
            End If
        End If
        rs_his_estructura.Close
    End If
    
    If ok_valor Then
        If Not EsNulo(htethasta) Then
            Call formato(CStr(htethasta), tfnro, v_tipovalor, v_valor)
        Else
            v_valor = ""
        End If
    End If
        
End Sub

Private Sub Depeve31(ByVal intnro As Integer, ByVal Empnro As Integer, ByVal ternro As Long, ByVal eventnro As Integer, ByVal topicnro As Integer, ByVal tfnro As Integer, ByRef v_tipovalor As Integer, ByRef v_valor As String, ByRef ok_valor As Boolean, ByRef msg_valor As String)
 Dim v_dec As String
 Dim intfhasta As Date
 Dim intfdesde As Date
 
 Dim rs_tercero As New ADODB.Recordset
 Dim rs_interfaz As New ADODB.Recordset
 Dim rs_his_estructura As New ADODB.Recordset
 
    v_tipovalor = 1
    ok_valor = True
    msg_valor = ""

    ' DTC_XXX - [JOB] - CONTRACT_TYPE
    StrSql = "SELECT * FROM tercero WHERE ternro = " & ternro
    OpenRecordset StrSql, rs_tercero
    If rs_tercero.EOF Then
        ok_valor = False
        msg_valor = "No se ha encontrado el Registro de Tercero para ternro " & CStr(ternro)
    End If
    rs_tercero.Close
    
    If ok_valor Then
        StrSql = "SELECT * FROM interfaz WHERE intnro = " & intnro
        OpenRecordset StrSql, rs_interfaz
        If rs_interfaz.EOF Then
            ok_valor = False
        Else
            intfhasta = rs_interfaz!intfhasta
            intfdesde = rs_interfaz!intfdesde
        End If
        rs_interfaz.Close
    End If
    
    If ok_valor Then
        StrSql = "SELECT estructura.estrcodext FROM his_estructura "
        StrSql = StrSql & " INNER JOIN estructura ON his_estructura.estrnro = estructura.estrnro "
        StrSql = StrSql & " WHERE ternro = " & ternro
        StrSql = StrSql & " AND his_estructura.tenro = 21"
        StrSql = StrSql & " AND htetdesde <= " & ConvFecha(intfhasta)
        StrSql = StrSql & " AND (htethasta >= " & ConvFecha(intfdesde) & " OR htethasta is null)"
        StrSql = StrSql & " ORDER BY htetdesde DESC"
        OpenRecordset StrSql, rs_his_estructura
        If rs_his_estructura.EOF Then
            v_dec = "1"
        Else
            If EsNulo(v_dec) Then
                v_dec = "1"
            Else
                v_dec = CStr(CDbl(rs_his_estructura!estrcodext) / CDbl(100))
            End If
        End If
        
        Call formato(CStr(v_dec), tfnro, v_tipovalor, v_valor)
        
        rs_his_estructura.Close
        
    End If
    
End Sub

Private Sub Depeve32(ByVal intnro As Integer, ByVal Empnro As Integer, ByVal ternro As Long, ByVal eventnro As Integer, ByVal topicnro As Integer, ByVal tfnro As Integer, ByRef v_tipovalor As Integer, ByRef v_valor As String, ByRef ok_valor As Boolean, ByRef msg_valor As String)
 Dim v_dec As String
 Dim intfhasta As Date
 Dim intfdesde As Date
 
 Dim rs_tercero As New ADODB.Recordset
 Dim rs_interfaz As New ADODB.Recordset
 Dim rs_his_estructura As New ADODB.Recordset
 
    v_tipovalor = 1
    ok_valor = True
    msg_valor = ""

    ' DTC_XXX - [JOB] - CONTRACT_TYPE
    StrSql = "SELECT * FROM tercero WHERE ternro = " & ternro
    OpenRecordset StrSql, rs_tercero
    If rs_tercero.EOF Then
        ok_valor = False
        msg_valor = "No se ha encontrado el Registro de Tercero para ternro " & CStr(ternro)
    End If
    rs_tercero.Close
    
    If ok_valor Then
        StrSql = "SELECT * FROM interfaz WHERE intnro = " & intnro
        OpenRecordset StrSql, rs_interfaz
        If rs_interfaz.EOF Then
            ok_valor = False
        Else
            intfhasta = rs_interfaz!intfhasta
            intfdesde = rs_interfaz!intfdesde
        End If
        rs_interfaz.Close
    End If
    
    If ok_valor Then
        StrSql = "SELECT formaliq.foliamposdo FROM his_estructura "
        StrSql = StrSql & " INNER JOIN formaliq ON his_estructura.estrnro = formaliq.estrnro "
        StrSql = StrSql & " WHERE ternro = " & ternro
        StrSql = StrSql & " AND his_estructura.tenro = 22"
        StrSql = StrSql & " AND htetdesde <= " & ConvFecha(intfhasta)
        StrSql = StrSql & " AND (htethasta >= " & ConvFecha(intfdesde) & " OR htethasta is null)"
        StrSql = StrSql & " ORDER BY htetdesde DESC"
        OpenRecordset StrSql, rs_his_estructura
    
        If rs_his_estructura.EOF Then
            v_dec = "1"
        Else
            If EsNulo(v_dec) Then
                v_dec = "1"
            Else
                v_dec = CStr(CDbl(rs_his_estructura!estrcodext) / CDbl(100))
            End If
        End If
        
        rs_his_estructura.Close
        
        Call formato(CStr(v_dec), tfnro, v_tipovalor, v_valor)
    End If
    
        
End Sub
                    
Private Sub Depeve33(ByVal intnro As Integer, ByVal Empnro As Integer, ByVal ternro As Long, ByVal eventnro As Integer, ByVal topicnro As Integer, ByVal tfnro As Integer, ByRef v_tipovalor As Integer, ByRef v_valor As String, ByRef ok_valor As Boolean, ByRef msg_valor As String)
 Dim codetosend As String
 Dim OK As Boolean
 
 Dim rs_emp_attrib_value As New ADODB.Recordset
 Dim rs_periodo As New ADODB.Recordset
 Dim rs_acu_mes As New ADODB.Recordset
 Dim rs_topic_field As New ADODB.Recordset
 
    ' Buscar si el empleado tiene asociado un valor con la entidad 15
    StrSql = "SELECT entity_value.codetosend FROM emp_attrib_value "
    StrSql = StrSql & " INNER JOIN entity_value ON emp_attrib_value.entvalnro = entity_value.entvalnro "
    StrSql = StrSql & " WHERE emp_attrib_value.ternro = " & ternro
    StrSql = StrSql & " AND entity_value.entnro = 15 "
'    StrSql = StrSql & " AND emp_attrib_value.empnro = " & empnro

    OpenRecordset StrSql, rs_emp_attrib_value
    
    If Not rs_emp_attrib_value.EOF Then
        codetosend = rs_emp_attrib_value!codetosend
        
        ' Buscar periodo de liq. Asociado.
        StrSql = "SELECT pliqanio, pliqmes  FROM interfaz "
        StrSql = StrSql & " INNER JOIN periodo ON interfaz.pliqnro = periodo.pliqnro "
        StrSql = StrSql & " WHERE interfaz.intnro = " & intnro
        
        OpenRecordset StrSql, rs_periodo
        
        If Not rs_periodo.EOF Then
            
            StrSql = "SELECT topic_field.entmapnro FROM topic_field WHERE topic_field.tfnro = " & tfnro
            OpenRecordset StrSql, rs_topic_field
            
            If Not rs_topic_field.EOF Then
                
                StrSql = "SELECT acu_mes.acunro, acu_mes.ammonto, entity_value.codetosend, entity_value.aditfield3 FROM acu_mes "
                StrSql = StrSql & " INNER JOIN entity_map_value ON acu_mes.acunro = entity_map_value.domvalor AND entity_map_value.entmapnro = " & rs_topic_field!entmapnro
                StrSql = StrSql & " INNER JOIN entity_value ON entity_map_value.imgvalor = entity_value.entvalnro "
                StrSql = StrSql & " WHERE "
                StrSql = StrSql & " acu_mes.ternro = " & ternro
                StrSql = StrSql & " AND acu_mes.amanio = " & rs_periodo!pliqanio
                StrSql = StrSql & " AND acu_mes.ammes = " & rs_periodo!pliqmes
                'StrSql = StrSql & " AND entity_map_value.empnro = " & empnro
                
                OpenRecordset StrSql, rs_acu_mes
                
                Do Until rs_acu_mes.EOF
                    If InStr(1, rs_acu_mes!aditfield3, codetosend) > 0 Then
                        v_valor = rs_acu_mes!codetosend
                        If rs_acu_mes!ammonto <> 0 Then
                            Call regeve01(intnro, Empnro, ternro, eventnro, topicnro, tfnro, 0, v_valor, OK)
                            
                            v_tipovalor = 3
                        End If
                    End If
                    rs_acu_mes.MoveNext
                Loop
                rs_acu_mes.Close
                
                If v_tipovalor <> 3 Then
                    v_tipovalor = 4
                End If
                
            End If
            rs_topic_field.Close
            
        End If
        
        rs_periodo.Close
        
    End If
    rs_emp_attrib_value.Close
    
End Sub
Private Sub Depeve34(ByVal intnro As Integer, ByVal Empnro As Integer, ByVal ternro As Long, ByVal eventnro As Integer, ByVal topicnro As Integer, ByVal tfnro As Integer, ByRef v_tipovalor As Integer, ByRef v_valor As String, ByRef ok_valor As Boolean, ByRef msg_valor As String)
 Dim codetosend As String
 Dim OK As Boolean
 Dim Existe1 As Boolean
 Dim v1 As Integer
 Dim vmult As String
 Dim vdecimal As Double
 
 Dim rs_emp_attrib_value As New ADODB.Recordset
 Dim rs_periodo As New ADODB.Recordset
 Dim rs_acu_mes As New ADODB.Recordset
 Dim rs_topic_field As New ADODB.Recordset
 Dim rs_intconfgen As New ADODB.Recordset
 Dim rs_entity_value As New ADODB.Recordset
 Dim rs_field_value As New ADODB.Recordset
 
    ' Buscar si el empleado tiene asociado un valor con la entidad 15
    StrSql = "SELECT entity_value.codetosend FROM emp_attrib_value "
    StrSql = StrSql & " INNER JOIN entity_value ON emp_attrib_value.entvalnro = entity_value.entvalnro "
    StrSql = StrSql & " WHERE emp_attrib_value.ternro = " & ternro
    StrSql = StrSql & " AND entity_value.entnro = 15 "
'    StrSql = StrSql & " AND emp_attrib_value.empnro = " & empnro

    OpenRecordset StrSql, rs_emp_attrib_value
    
    If Not rs_emp_attrib_value.EOF Then
        codetosend = rs_emp_attrib_value!codetosend
        
        ' Buscar periodo de liq. Asociado.
        StrSql = "SELECT pliqanio, pliqmes  FROM interfaz "
        StrSql = StrSql & " INNER JOIN periodo ON interfaz.pliqnro = periodo.pliqnro "
        StrSql = StrSql & " WHERE interfaz.intnro = " & intnro
        
        OpenRecordset StrSql, rs_periodo
        
        If Not rs_periodo.EOF Then
            
            StrSql = "SELECT topic_field.entmapnro FROM topic_field WHERE topic_field.tfnro = " & tfnro
            OpenRecordset StrSql, rs_topic_field
            
            If Not rs_topic_field.EOF Then
                'INICIO {multi.i vmult}
                ' Buscar entidad que tiene los valores multiplicadores
                StrSql = "SELECT * FROM intconfgen INNER JOIN entity ON intconfgen.entnro = entity.entnro "
                OpenRecordset StrSql, rs_intconfgen
                
                If rs_intconfgen.EOF Then
                    ok_valor = False
                    msg_valor = "No esta conf. Entidad asociada."
                    Exit Sub
                Else
                    ' Buscar valor asociado
                
                    StrSql = "SELECT emp_attrib_value.entvalnro FROM emp_attrib_value "
                    StrSql = StrSql & " INNER JOIN entity_value ON emp_attrib_value.entvalnro = entity_value.entvalnro "
                    StrSql = StrSql & " WHERE emp_attrib_value.ternro = " & ternro & " AND entity_value.entnro = " & rs_intconfgen!entnro
                    OpenRecordset StrSql, rs_emp_attrib_value
                    
                    Existe1 = False
                    If rs_emp_attrib_value.EOF Then
                        ok_valor = False
                        msg_valor = "No esta conf. Manual para campo " + CStr(tfnro) + ". ternro " + CStr(ternro)
                        Exit Sub
                    Else
                        Existe1 = True
                        v1 = rs_emp_attrib_value!entvalnro
                    End If
                    rs_emp_attrib_value.Close
                End If
                
                ' Buscar valor en su entidad correspondiente
                StrSql = "SELECT codetosend FROM entity_value "
                StrSql = StrSql & " WHERE entity_value.entvalnro = " & v1
                OpenRecordset StrSql, rs_entity_value
                
                If Not rs_entity_value.EOF Then
                    vmult = rs_entity_value!codetosend
                Else
                    ok_valor = False
                    msg_valor = "No existe valor Conf. Manual para campo  " + CStr(tfnro) + ". Valor " + CStr(v1)
                    Exit Sub
                End If
                rs_entity_value.Close
                
                'FIN {multi.i vmult}

                StrSql = "SELECT acu_mes.acunro, acu_mes.ammonto, entity_value.codetosend, entity_value.aditfield3 FROM acu_mes "
                StrSql = StrSql & " INNER JOIN entity_map_value ON acu_mes.acunro = entity_map_value.domvalor AND entity_map_value.entmapnro = " & rs_topic_field!entmapnro
                StrSql = StrSql & " INNER JOIN entity_value ON entity_map_value.imgvalor = entity_value.entvalnro "
                StrSql = StrSql & " WHERE "
                StrSql = StrSql & " acu_mes.ternro = " & ternro
                StrSql = StrSql & " AND acu_mes.amanio = " & rs_periodo!pliqanio
                StrSql = StrSql & " AND acu_mes.ammes = " & rs_periodo!pliqmes
                'StrSql = StrSql & " AND entity_map_value.empnro = " & empnro
                
                OpenRecordset StrSql, rs_acu_mes
                
                Do Until rs_acu_mes.EOF
                    If InStr(1, rs_acu_mes!aditfield3, codetosend) > 0 Then
                        v_valor = ""
                        vdecimal = CDbl(rs_acu_mes!ammonto) * CDbl(Replace(vmult, ",", "."))
                        
                        Call formato(CStr(vdecimal), tfnro, v_tipovalor, v_valor)
                
                        ' Registrar valor para Int-Ter-Empresa-Topico-Campo
                        Call regeve01(intnro, Empnro, ternro, eventnro, topicnro, tfnro, 0, v_valor, OK)
                        
                        v_tipovalor = 3
                    End If
                    
                    rs_acu_mes.MoveNext
                    
                Loop
                rs_acu_mes.Close
                
                If v_tipovalor <> 3 Then
                    v_tipovalor = 4
                End If
                
            End If
            rs_topic_field.Close
            
        End If
        
        rs_periodo.Close
        
    End If
    'rs_emp_attrib_value.Close
    
End Sub
                    
Private Sub Depeve35(ByVal intnro As Integer, ByVal Empnro As Integer, ByVal ternro As Long, ByVal eventnro As Integer, ByVal topicnro As Integer, ByVal tfnro As Integer, ByRef v_tipovalor As Integer, ByRef v_valor As String, ByRef ok_valor As Boolean, ByRef msg_valor As String)
 Dim codetosend As String
 Dim OK As Boolean
 
 Dim rs_emp_attrib_value As New ADODB.Recordset
 Dim rs_periodo As New ADODB.Recordset
 Dim rs_acu_mes As New ADODB.Recordset
 Dim rs_topic_field As New ADODB.Recordset
 
    ' Buscar si el empleado tiene asociado un valor con la entidad 15
    StrSql = "SELECT entity_value.codetosend FROM emp_attrib_value "
    StrSql = StrSql & " INNER JOIN entity_value ON emp_attrib_value.entvalnro = entity_value.entvalnro "
    StrSql = StrSql & " WHERE emp_attrib_value.ternro = " & ternro
    StrSql = StrSql & " AND entity_value.entnro = 15 "
'    StrSql = StrSql & " AND emp_attrib_value.empnro = " & empnro

    OpenRecordset StrSql, rs_emp_attrib_value
    
    If Not rs_emp_attrib_value.EOF Then
        codetosend = rs_emp_attrib_value!codetosend
        
        ' Buscar periodo de liq. Asociado.
        StrSql = "SELECT pliqanio, pliqmes  FROM interfaz "
        StrSql = StrSql & " INNER JOIN periodo ON interfaz.pliqnro = periodo.pliqnro "
        StrSql = StrSql & " WHERE interfaz.intnro = " & intnro
        
        OpenRecordset StrSql, rs_periodo
        
        If Not rs_periodo.EOF Then
            
            StrSql = "SELECT topic_field.entmapnro FROM topic_field WHERE topic_field.tfnro = " & tfnro
            OpenRecordset StrSql, rs_topic_field
            
            If Not rs_topic_field.EOF Then
                
                StrSql = "SELECT acu_mes.acunro, acu_mes.ammonto, entity_value.codetosend, entity_value.aditfield3 FROM acu_mes "
                StrSql = StrSql & " INNER JOIN entity_map_value ON acu_mes.acunro = entity_map_value.domvalor AND entity_map_value.entmapnro = " & rs_topic_field!entmapnro
                StrSql = StrSql & " INNER JOIN entity_value ON entity_map_value.imgvalor = entity_value.entvalnro "
                StrSql = StrSql & " WHERE "
                StrSql = StrSql & " acu_mes.ternro = " & ternro
                StrSql = StrSql & " AND acu_mes.amanio = " & rs_periodo!pliqanio
                StrSql = StrSql & " AND acu_mes.ammes = " & rs_periodo!pliqmes
                'StrSql = StrSql & " AND entity_map_value.empnro = " & empnro
                
                OpenRecordset StrSql, rs_acu_mes
                
                Do Until rs_acu_mes.EOF
                    v_valor = rs_acu_mes!codetosend
                    If rs_acu_mes!ammonto <> 0 Then
                        Call regeve01(intnro, Empnro, ternro, eventnro, topicnro, tfnro, 0, v_valor, OK)
                        
                        v_tipovalor = 3
                    End If
                    rs_acu_mes.MoveNext
                    
                Loop
                rs_acu_mes.Close
                
                If v_tipovalor <> 3 Then
                    v_tipovalor = 4
                End If
                
            End If
            rs_topic_field.Close
            
        End If
        
        rs_periodo.Close
        
    End If
    rs_emp_attrib_value.Close
    
End Sub

Private Sub Depeve36(ByVal intnro As Integer, ByVal Empnro As Integer, ByVal ternro As Long, ByVal eventnro As Integer, ByVal topicnro As Integer, ByVal tfnro As Integer, ByRef v_tipovalor As Integer, ByRef v_valor As String, ByRef ok_valor As Boolean, ByRef msg_valor As String)
 Dim codetosend As String
 Dim OK As Boolean
 Dim Existe1 As Boolean
 Dim v1 As Integer
 Dim vmult As String
 Dim vdecimal As Double
 
 Dim rs_emp_attrib_value As New ADODB.Recordset
 Dim rs_periodo As New ADODB.Recordset
 Dim rs_acu_mes As New ADODB.Recordset
 Dim rs_topic_field As New ADODB.Recordset
 Dim rs_intconfgen As New ADODB.Recordset
 Dim rs_entity_value As New ADODB.Recordset
 Dim rs_field_value As New ADODB.Recordset
 
    ' Buscar si el empleado tiene asociado un valor con la entidad 15
    StrSql = "SELECT entity_value.codetosend FROM emp_attrib_value "
    StrSql = StrSql & " INNER JOIN entity_value ON emp_attrib_value.entvalnro = entity_value.entvalnro "
    StrSql = StrSql & " WHERE emp_attrib_value.ternro = " & ternro
    StrSql = StrSql & " AND entity_value.entnro = 15 "
'    StrSql = StrSql & " AND emp_attrib_value.empnro = " & empnro

    OpenRecordset StrSql, rs_emp_attrib_value
    
    If Not rs_emp_attrib_value.EOF Then
        codetosend = rs_emp_attrib_value!codetosend
        
        ' Buscar periodo de liq. Asociado.
        StrSql = "SELECT pliqanio, pliqmes  FROM interfaz "
        StrSql = StrSql & " INNER JOIN periodo ON interfaz.pliqnro = periodo.pliqnro "
        StrSql = StrSql & " WHERE interfaz.intnro = " & intnro
        
        OpenRecordset StrSql, rs_periodo
        
        If Not rs_periodo.EOF Then
            
            StrSql = "SELECT topic_field.entmapnro FROM topic_field WHERE topic_field.tfnro = " & tfnro
            OpenRecordset StrSql, rs_topic_field
            
            If Not rs_topic_field.EOF Then
                'INICIO {multi.i vmult}
                ' Buscar entidad que tiene los valores multiplicadores
                StrSql = "SELECT * FROM intconfgen INNER JOIN entity ON intconfgen.entnro = entity.entnro "
                OpenRecordset StrSql, rs_intconfgen
                
                If rs_intconfgen.EOF Then
                    ok_valor = False
                    msg_valor = "No esta conf. Entidad asociada."
                    Exit Sub
                Else
                    ' Buscar valor asociado
                
                    StrSql = "SELECT emp_attrib_value.entvalnro FROM emp_attrib_value "
                    StrSql = StrSql & " INNER JOIN entity_value ON emp_attrib_value.entvalnro = entity_value.entvalnro "
                    StrSql = StrSql & " WHERE emp_attrib_value.ternro = " & ternro & " AND entity_value.entnro = " & rs_intconfgen!entnro
                    OpenRecordset StrSql, rs_emp_attrib_value
                    
                    Existe1 = False
                    If rs_emp_attrib_value.EOF Then
                        ok_valor = False
                        msg_valor = "No esta conf. Manual para campo " + CStr(tfnro) + ". ternro " + CStr(ternro)
                        Exit Sub
                    Else
                        Existe1 = True
                        v1 = rs_emp_attrib_value!entvalnro
                    End If
                    rs_emp_attrib_value.Close
                End If
                
                ' Buscar valor en su entidad correspondiente
                StrSql = "SELECT codetosend FROM entity_value "
                StrSql = StrSql & " WHERE entity_value.entvalnro = " & v1
                OpenRecordset StrSql, rs_entity_value
                
                If Not rs_entity_value.EOF Then
                    vmult = rs_entity_value!codetosend
                Else
                    ok_valor = False
                    msg_valor = "No existe valor Conf. Manual para campo  " + CStr(tfnro) + ". Valor " + CStr(v1)
                    Exit Sub
                End If
                rs_entity_value.Close
                
                'FIN {multi.i vmult}

                StrSql = "SELECT acu_mes.acunro, acu_mes.ammonto, entity_value.codetosend, entity_value.aditfield3 FROM acu_mes "
                StrSql = StrSql & " INNER JOIN entity_map_value ON acu_mes.acunro = entity_map_value.domvalor AND entity_map_value.entmapnro = " & rs_topic_field!entmapnro
                StrSql = StrSql & " INNER JOIN entity_value ON entity_map_value.imgvalor = entity_value.entvalnro "
                StrSql = StrSql & " WHERE "
                StrSql = StrSql & " acu_mes.ternro = " & ternro
                StrSql = StrSql & " AND acu_mes.amanio = " & rs_periodo!pliqanio
                StrSql = StrSql & " AND acu_mes.ammes = " & rs_periodo!pliqmes
                'StrSql = StrSql & " AND entity_map_value.empnro = " & empnro
                
                OpenRecordset StrSql, rs_acu_mes
                
                Do Until rs_acu_mes.EOF
                    v_valor = ""
                    vdecimal = CDbl(rs_acu_mes!ammonto)
                    
                    Call formato(CStr(vdecimal), tfnro, v_tipovalor, v_valor)
            
                    ' Registrar valor para Int-Ter-Empresa-Topico-Campo
                    Call regeve01(intnro, Empnro, ternro, eventnro, topicnro, tfnro, 0, v_valor, OK)
                    
                    v_tipovalor = 3
                    
                    rs_acu_mes.MoveNext
                    
                Loop
                rs_acu_mes.Close
                
                If v_tipovalor <> 3 Then
                    v_tipovalor = 4
                End If
                
            End If
            rs_topic_field.Close
            
        End If
        
        rs_periodo.Close
        
    End If
    'rs_emp_attrib_value.Close
    
End Sub

Private Sub Depeve41(ByVal intnro As Integer, ByVal Empnro As Integer, ByVal ternro As Long, ByVal eventnro As Integer, ByVal topicnro As Integer, ByVal tfnro As Integer, ByRef v_tipovalor As Integer, ByRef v_valor As String, ByRef ok_valor As Boolean, ByRef msg_valor As String)
 Dim rs_periodo As New ADODB.Recordset
 
    v_tipovalor = 1
    ok_valor = True
    msg_valor = ""
    
    ' Buscar periodo para la empresa actual
    StrSql = "SELECT * FROM periodo "
    StrSql = StrSql & " INNER JOIN interfaz ON periodo.pliqnro = interfaz.pliqnro WHERE interfaz.intnro = " & intnro
    OpenRecordset StrSql, rs_periodo
    If rs_periodo.EOF Then
        ok_valor = False
        msg_valor = "No se ha encontrado los datos del Período para la interface " & CStr(intnro)
    Else
        Call formato(CStr(rs_periodo!pliqdesde), tfnro, v_tipovalor, v_valor)
    End If
    
    rs_periodo.Close
    
End Sub

Private Sub Depeve42(ByVal intnro As Integer, ByVal Empnro As Integer, ByVal ternro As Long, ByVal eventnro As Integer, ByVal topicnro As Integer, ByVal tfnro As Integer, ByRef v_tipovalor As Integer, ByRef v_valor As String, ByRef ok_valor As Boolean, ByRef msg_valor As String)
 Dim codetosend As String
 Dim OK As Boolean
 Dim Existe1 As Boolean
 Dim v1 As Integer
 Dim vmult As String
 Dim vdecimal As Double

 Dim rs_emp_attrib_value As New ADODB.Recordset
 Dim rs_periodo As New ADODB.Recordset
 Dim rs_acu_mes As New ADODB.Recordset
 Dim rs_topic_field As New ADODB.Recordset
 Dim rs_intconfgen As New ADODB.Recordset
 Dim rs_entity_value As New ADODB.Recordset
 Dim rs_b_interfaz As New ADODB.Recordset
 Dim rs_field_value As New ADODB.Recordset
 
    ' Buscar si el empleado tiene asociado un valor con la entidad 15
    StrSql = "SELECT entity_value.codetosend FROM emp_attrib_value "
    StrSql = StrSql & " INNER JOIN entity_value ON emp_attrib_value.entvalnro = entity_value.entvalnro "
    StrSql = StrSql & " WHERE emp_attrib_value.ternro = " & ternro
    StrSql = StrSql & " AND entity_value.entnro = 15 "
'    StrSql = StrSql & " AND emp_attrib_value.empnro = " & empnro

    OpenRecordset StrSql, rs_emp_attrib_value
    
    If Not rs_emp_attrib_value.EOF Then
        codetosend = rs_emp_attrib_value!codetosend
        
        ' Buscar periodo de liq. Asociado.
        StrSql = "SELECT pliqanio, pliqmes  FROM interfaz "
        StrSql = StrSql & " INNER JOIN periodo ON interfaz.pliqnro = periodo.pliqnro "
        StrSql = StrSql & " WHERE interfaz.intnro = " & intnro
        
        OpenRecordset StrSql, rs_periodo
        
        If Not rs_periodo.EOF Then
            
            StrSql = "SELECT topic_field.entmapnro FROM topic_field WHERE topic_field.tfnro = " & tfnro
            OpenRecordset StrSql, rs_topic_field
            
            If Not rs_topic_field.EOF Then
                'INICIO {multi.i vmult}
                ' Buscar entidad que tiene los valores multiplicadores
                StrSql = "SELECT * FROM intconfgen INNER JOIN entity ON intconfgen.entnro = entity.entnro "
                OpenRecordset StrSql, rs_intconfgen
                
                If rs_intconfgen.EOF Then
                    ok_valor = False
                    msg_valor = "No esta conf. Entidad asociada."
                    Exit Sub
                Else
                    ' Buscar valor asociado
                
                    StrSql = "SELECT emp_attrib_value.entvalnro FROM emp_attrib_value "
                    StrSql = StrSql & " INNER JOIN entity_value ON emp_attrib_value.entvalnro = entity_value.entvalnro "
                    StrSql = StrSql & " WHERE emp_attrib_value.ternro = " & ternro & " AND entity_value.entnro = " & rs_intconfgen!entnro
                    OpenRecordset StrSql, rs_emp_attrib_value
                    
                    Existe1 = False
                    If rs_emp_attrib_value.EOF Then
                        ok_valor = False
                        msg_valor = "No esta conf. Manual para campo " + CStr(tfnro) + ". ternro " + CStr(ternro)
                        Exit Sub
                    Else
                        Existe1 = True
                        v1 = rs_emp_attrib_value!entvalnro
                    End If
                    rs_emp_attrib_value.Close
                End If
                
                ' Buscar valor en su entidad correspondiente
                StrSql = "SELECT codetosend FROM entity_value "
                StrSql = StrSql & " WHERE entity_value.entvalnro = " & v1
                OpenRecordset StrSql, rs_entity_value
                
                If Not rs_entity_value.EOF Then
                    vmult = rs_entity_value!codetosend
                Else
                    ok_valor = False
                    msg_valor = "No existe valor Conf. Manual para campo  " + CStr(tfnro) + ". Valor " + CStr(v1)
                    Exit Sub
                End If
                rs_entity_value.Close
                
                'FIN {multi.i vmult}

                StrSql = "SELECT acu_mes.acunro, acu_mes.ammonto, entity_value.codetosend, entity_value.aditfield3 FROM acu_mes "
                StrSql = StrSql & " INNER JOIN entity_map_value ON acu_mes.acunro = entity_map_value.domvalor AND entity_map_value.entmapnro = " & rs_topic_field!entmapnro
                StrSql = StrSql & " INNER JOIN entity_value ON entity_map_value.imgvalor = entity_value.entvalnro "
                StrSql = StrSql & " WHERE "
                StrSql = StrSql & " acu_mes.ternro = " & ternro
                StrSql = StrSql & " AND acu_mes.amanio = " & rs_periodo!pliqanio
                StrSql = StrSql & " AND acu_mes.ammes = " & rs_periodo!pliqmes
                'StrSql = StrSql & " AND entity_map_value.empnro = " & empnro
                
                OpenRecordset StrSql, rs_acu_mes
                
                Do Until rs_acu_mes.EOF
                    v_valor = ""
                    vdecimal = CDbl(rs_acu_mes!ammonto)
                    
                    ' Buscar valor registrado para el mismo topico en la
                    ' ultima interfaz aceptada
                    StrSql = "SELECT * FROM interfaz WHERE intnro < " & intnro & " AND intok = -1 ORDER BY intnro DESC"
                    OpenRecordset StrSql, rs_b_interfaz
                    
                    If Not rs_b_interfaz.EOF Then
                        StrSql = "SELECT * FROM field_value WHERE field_value.archivado = -1 "
                        StrSql = StrSql & " AND field_value.intnro = " & rs_b_interfaz!intnro
                        StrSql = StrSql & " AND field_value.ternro = " & ternro
                        StrSql = StrSql & " AND field_value.tfnro = " & tfnro
                        StrSql = StrSql & " AND field_value.topicnro = " & topicnro
                        'StrSql = StrSql & "FIELD_VALUE.empnro = nroemp
                        OpenRecordset StrSql, rs_field_value
                                                
                        If Not rs_field_value.EOF Then
                            vdecimal = vdecimal + CDbl(rs_field_value!Valor)
                        End If
                        
                        Call formato(CStr(vdecimal), tfnro, v_tipovalor, v_valor)
            
                        ' Registrar valor para Int-Ter-Empresa-Topico-Campo
                        Call regeve01(intnro, Empnro, ternro, eventnro, topicnro, tfnro, 0, v_valor, OK)
                            
                        v_tipovalor = 3
                
                    End If
                    rs_acu_mes.MoveNext
                    
                Loop
                rs_acu_mes.Close
                
                If v_tipovalor <> 3 Then
                    v_tipovalor = 4
                End If
                
            End If
            rs_topic_field.Close
            
        End If
        
        rs_periodo.Close
        
    End If
    'rs_emp_attrib_value.Close
    
End Sub

Private Sub Depeve43(ByVal intnro As Integer, ByVal Empnro As Integer, ByVal ternro As Long, ByVal eventnro As Integer, ByVal topicnro As Integer, ByVal tfnro As Integer, ByRef v_tipovalor As Integer, ByRef v_valor As String, ByRef ok_valor As Boolean, ByRef msg_valor As String)
 Dim intfhasta As Date
 Dim mindiaslic As Integer
 Dim emp_licnro As Integer
 Dim elfechahasta As Date
 Dim elfechadesde As Date
 Dim v_fecha As Date
 Dim tdnro As Integer
 
 Dim rs_empleado As New ADODB.Recordset
 Dim rs_interfaz As New ADODB.Recordset
 Dim rs_intconfgen As New ADODB.Recordset
 Dim rs_emp_lic As New ADODB.Recordset
 Dim rs_b_Emp_Lic As New ADODB.Recordset
 
 ' v_tipovalor = 1 Caracter
 ' v_tipovalor = 2 Código interno de valor válido (bnp_entity_value.entvalnro)
 
    v_tipovalor = 1
    ok_valor = True
    msg_valor = ""
    
    ' Buscar el periodo de interfaz
    StrSql = "SELECT * FROM interfaz WHERE intnro = " & intnro
    OpenRecordset StrSql, rs_interfaz
    If rs_interfaz.EOF Then
        ok_valor = False
    Else
        intfhasta = rs_interfaz!intfhasta
    End If
    rs_interfaz.Close
    
    If ok_valor Then
        StrSql = "SELECT * FROM intconfgen"
        OpenRecordset StrSql, rs_intconfgen
        If rs_intconfgen.EOF Then
            ok_valor = False
        Else
            mindiaslic = rs_intconfgen!mindiaslic
        End If
        rs_intconfgen.Close
    End If
                
    If ok_valor Then
        ' Buscar licencia del empleado
        StrSql = "SELECT * FROM emp_lic "
        StrSql = StrSql & " WHERE emp_lic.empleado = " & ternro
        StrSql = StrSql & " AND elfechadesde <= " & ConvFecha(intfhasta) & " AND elfechahasta >= " & ConvFecha(intfhasta)
        OpenRecordset StrSql, rs_emp_lic
        If rs_emp_lic.EOF Then
            ok_valor = False
        Else
            emp_licnro = rs_emp_lic!emp_licnro
            elfechahasta = rs_emp_lic!elfechahasta
            elfechadesde = rs_emp_lic!elfechadesde
            tdnro = rs_emp_lic!tdnro
        End If
        rs_emp_lic.Close
    End If
        
    If ok_valor Then
        If DateDiff("d", elfechadesde, elfechahasta) >= mindiaslic Then
            v_fecha = elfechadesde
        Else
            ' Buscar otra licencia anterior que tenga las siguientes condiciones:
            '   1.- Debe ser de menos de 28 dias corridos
            '   2.- Debe ser del mismo tipo
            '   3.- Debe ser continua
            If Weekday(elfechadesde) = 2 Then
                ' Si elfechadesde es Lunes, debo considerar que la licencia termine un viernes, sabado o domingo
                StrSql = "SELECT * FROM emp_lic "
                StrSql = StrSql & " WHERE empleado = " & ternro
                StrSql = StrSql & " AND ((elfechahasta = " & DateAdd("d", -3, elfechadesde) ' viernes
                StrSql = StrSql & " ) OR (elfechahasta = " & DateAdd("d", -2, elfechadesde) ' sabado
                StrSql = StrSql & " ) OR (elfechahasta = " & DateAdd("d", -1, elfechadesde) ' domingo
                StrSql = StrSql & " )) AND tdnro = " & tdnro
            Else
                StrSql = "SELECT * FROM emp_lic "
                StrSql = StrSql & " WHERE empleado = " & ternro
                StrSql = StrSql & " AND (elfechahasta = " & DateAdd("d", -1, elfechadesde)
                StrSql = StrSql & " ) AND tdnro = " & tdnro
            End If
            
            OpenRecordset StrSql, rs_b_Emp_Lic
            
            If Not rs_b_Emp_Lic.EOF Then
                If ((DateDiff("d", rs_b_Emp_Lic!elfechadesde, rs_b_Emp_Lic!elfechahasta) < mindiaslic) And (DateDiff("d", rs_b_Emp_Lic!elfechadesde, elfechahasta) >= mindiaslic)) Then
                    v_fecha = rs_b_Emp_Lic!elfechadesde
                Else
                    ok_valor = False
                End If
            Else
                ok_valor = False
            End If
            
            rs_b_Emp_Lic.Close
            
        End If
    End If
    
    If ok_valor Then
        Call formato(CStr(v_fecha), tfnro, v_tipovalor, v_valor)
    End If

End Sub

Private Sub Depeve44(ByVal intnro As Integer, ByVal Empnro As Integer, ByVal ternro As Long, ByVal eventnro As Integer, ByVal topicnro As Integer, ByVal tfnro As Integer, ByRef v_tipovalor As Integer, ByRef v_valor As String, ByRef ok_valor As Boolean, ByRef msg_valor As String)
 Dim intfhasta As Date
 Dim mindiaslic As Integer
 Dim emp_licnro As Integer
 Dim elfechahasta As Date
 Dim elfechadesde As Date
 Dim tdnro As Integer
 Dim imgvalor As Integer
 
 Dim rs_interfaz As New ADODB.Recordset
 Dim rs_intconfgen As New ADODB.Recordset
 Dim rs_emp_lic As New ADODB.Recordset
 Dim rs_entity_map_value As New ADODB.Recordset
 Dim rs_entity_value As New ADODB.Recordset
 
 ' v_tipovalor = 1 Caracter
 ' v_tipovalor = 2 Código interno de valor válido (bnp_entity_value.entvalnro)
 
    v_tipovalor = 1
    ok_valor = True
    msg_valor = ""
    
    ' Buscar el periodo de interfaz
    StrSql = "SELECT * FROM interfaz WHERE intnro = " & intnro
    OpenRecordset StrSql, rs_interfaz
    If rs_interfaz.EOF Then
        ok_valor = False
    Else
        intfhasta = rs_interfaz!intfhasta
    End If
    rs_interfaz.Close
    
    If ok_valor Then
        StrSql = "SELECT * FROM intconfgen"
        OpenRecordset StrSql, rs_intconfgen
        If rs_intconfgen.EOF Then
            ok_valor = False
        Else
            mindiaslic = rs_intconfgen!mindiaslic
        End If
        rs_intconfgen.Close
    End If
                
    If ok_valor Then
        ' Buscar licencia del empleado
        StrSql = "SELECT * FROM emp_lic "
        StrSql = StrSql & " WHERE emp_lic.empleado = " & ternro
        StrSql = StrSql & " AND elfechadesde <= " & ConvFecha(intfhasta) & " AND elfechahasta >= " & ConvFecha(intfhasta)
        OpenRecordset StrSql, rs_emp_lic
        If rs_emp_lic.EOF Then
            ok_valor = False
        Else
            emp_licnro = rs_emp_lic!emp_licnro
            elfechahasta = rs_emp_lic!elfechahasta
            elfechadesde = rs_emp_lic!elfechadesde
            tdnro = rs_emp_lic!tdnro
        End If
        rs_emp_lic.Close
    End If
        
    If ok_valor Then
        StrSql = "SELECT imgvalor FROM entity_map_value "
        StrSql = StrSql & " INNER JOIN entity_map ON entity_map_value.entmapnro = entity_map.entmapnro"
        StrSql = StrSql & " WHERE entity_map_value.domvalor = " & tdnro & " AND entity_map.entmapnro = 11"
        OpenRecordset StrSql, rs_entity_map_value
        If rs_entity_map_value.EOF Then
            ok_valor = False
        Else
            imgvalor = rs_entity_map_value!imgvalor
        End If
        rs_entity_map_value.Close
    End If
    
    If ok_valor Then
        StrSql = "SELECT codetosend FROM entity_value "
        StrSql = StrSql & " WHERE entity_value.entvalnro = " & imgvalor
        OpenRecordset StrSql, rs_entity_value
        If rs_entity_value.EOF Then
            ok_valor = False
        Else
            v_valor = rs_entity_value!codetosend
        End If
        rs_entity_value.Close
    End If
    
End Sub

Private Sub Depeve45(ByVal intnro As Integer, ByVal Empnro As Integer, ByVal ternro As Long, ByVal eventnro As Integer, ByVal topicnro As Integer, ByVal tfnro As Integer, ByRef v_tipovalor As Integer, ByRef v_valor As String, ByRef ok_valor As Boolean, ByRef msg_valor As String)
 Dim intfhasta As Date
 Dim intfdesde As Date
 Dim mindiaslic As Integer
 Dim v_fecha As Date
 
 Dim rs_interfaz As New ADODB.Recordset
 Dim rs_intconfgen As New ADODB.Recordset
 Dim rs_emp_lic As New ADODB.Recordset
 
 ' The event "Return after Leave" (RFL) is used to record the end of a leave
 ' previously notified by an event "Paid Leave" (PLA) or " Unpaid Leave" (LOA).
 
 ' v_tipovalor = 1 Caracter
 ' v_tipovalor = 2 Código interno de valor válido (bnp_entity_value.entvalnro)
 
 ' Buscar una licencia de la persona que termine en este período de la interfaz
 
    v_tipovalor = 1
    ok_valor = True
    msg_valor = ""
    
    ' Buscar el periodo de interfaz
    StrSql = "SELECT * FROM interfaz WHERE intnro = " & intnro
    OpenRecordset StrSql, rs_interfaz
    If rs_interfaz.EOF Then
        ok_valor = False
    Else
        intfhasta = rs_interfaz!intfhasta
        intfdesde = rs_interfaz!intfdesde
    End If
    rs_interfaz.Close
    
    If ok_valor Then
        StrSql = "SELECT * FROM intconfgen"
        OpenRecordset StrSql, rs_intconfgen
        If rs_intconfgen.EOF Then
            ok_valor = False
        Else
            mindiaslic = rs_intconfgen!mindiaslic
        End If
        rs_intconfgen.Close
    End If
                
    If ok_valor Then
        ' Buscar licencia del empleado
        StrSql = "SELECT * FROM emp_lic "
        StrSql = StrSql & " WHERE emp_lic.empleado = " & ternro
        StrSql = StrSql & " AND elfechahasta >= " & ConvFecha(DateAdd("d", -1, intfdesde)) & " AND elfechahasta <= " & ConvFecha(DateAdd("d", -1, intfhasta))
        OpenRecordset StrSql, rs_emp_lic
        If rs_emp_lic.EOF Then
            ok_valor = False
        Else
            v_fecha = DateAdd("d", 1, rs_emp_lic!elfechahasta)
        End If
        rs_emp_lic.Close
    End If
    
    If ok_valor Then
        Call formato(CStr(v_fecha), tfnro, v_tipovalor, v_valor)
    End If
    
End Sub

Private Sub Depeve46(ByVal intnro As Integer, ByVal Empnro As Integer, ByVal ternro As Long, ByVal eventnro As Integer, ByVal topicnro As Integer, ByVal tfnro As Integer, ByRef v_tipovalor As Integer, ByRef v_valor As String, ByRef ok_valor As Boolean, ByRef msg_valor As String)
 Dim intfhasta As Date
 Dim intfdesde As Date
 Dim v_fecha As Date
 
 Dim rs_interfaz As New ADODB.Recordset
 Dim rs_fases As New ADODB.Recordset
 
 ' v_tipovalor = 1 Caracter
 ' v_tipovalor = 2 Código interno de valor válido (bnp_entity_value.entvalnro)
 
    v_tipovalor = 1
    ok_valor = True
    msg_valor = ""
    
    ' Buscar el periodo de interfaz
    StrSql = "SELECT * FROM interfaz WHERE intnro = " & intnro
    OpenRecordset StrSql, rs_interfaz
    If rs_interfaz.EOF Then
        ok_valor = False
    Else
        intfhasta = rs_interfaz!intfhasta
        intfdesde = rs_interfaz!intfdesde
    End If
    rs_interfaz.Close
    
    If ok_valor Then
        ' Buscar baja dentro del período de la interfaz
        StrSql = "SELECT * FROM fases "
        StrSql = StrSql & " WHERE fases.empleado = " & ternro
        StrSql = StrSql & " AND bajfec >= " & ConvFecha(intfdesde) & " AND bajfec <= " & ConvFecha(intfhasta)
        StrSql = StrSql & " AND estado = 0 ORDER BY altfec DESC"
        OpenRecordset StrSql, rs_fases
        If rs_fases.EOF Then
            ok_valor = False
        Else
            v_fecha = rs_fases!bajfec
        End If
        rs_fases.Close
    End If
    
    If ok_valor Then
        Call formato(CStr(v_fecha), tfnro, v_tipovalor, v_valor)
    End If
    
End Sub

Private Sub Depeve47(ByVal intnro As Integer, ByVal Empnro As Integer, ByVal ternro As Long, ByVal eventnro As Integer, ByVal topicnro As Integer, ByVal tfnro As Integer, ByRef v_tipovalor As Integer, ByRef v_valor As String, ByRef ok_valor As Boolean, ByRef msg_valor As String)
 Dim intfhasta As Date
 Dim intfdesde As Date
 Dim caunro As Integer
 Dim imgvalor As Integer
 
 Dim rs_interfaz As New ADODB.Recordset
 Dim rs_fases As New ADODB.Recordset
 Dim rs_entity_map_value As New ADODB.Recordset
 Dim rs_entity_value As New ADODB.Recordset
 
 ' v_tipovalor = 1 Caracter
 ' v_tipovalor = 2 Código interno de valor válido (bnp_entity_value.entvalnro)
 
    v_tipovalor = 1
    ok_valor = True
    msg_valor = ""
    
    ' Buscar el periodo de interfaz
    StrSql = "SELECT * FROM interfaz WHERE intnro = " & intnro
    OpenRecordset StrSql, rs_interfaz
    If rs_interfaz.EOF Then
        ok_valor = False
    Else
        intfhasta = rs_interfaz!intfhasta
        intfdesde = rs_interfaz!intfdesde
    End If
    rs_interfaz.Close
    
    If ok_valor Then
        ' Buscar baja dentro del período de la interfaz
        StrSql = "SELECT caunro FROM fases "
        StrSql = StrSql & " WHERE fases.empleado = " & ternro
        StrSql = StrSql & " AND bajfec >= " & ConvFecha(intfdesde) & " AND bajfec <= " & ConvFecha(intfhasta)
        StrSql = StrSql & " AND estado = 0 ORDER BY altfec DESC"
        OpenRecordset StrSql, rs_fases
        If rs_fases.EOF Then
            ok_valor = False
        Else
            caunro = rs_fases!caunro
        End If
        rs_fases.Close
    End If
    
    If ok_valor Then
        StrSql = "SELECT imgvalor FROM entity_map_value "
        StrSql = StrSql & " INNER JOIN entity_map ON entity_map_value.entmapnro = entity_map.entmapnro"
        StrSql = StrSql & " WHERE entity_map_value.domvalor = " & caunro & " AND entity_map.entmapnro = 13"
        OpenRecordset StrSql, rs_entity_map_value
        If rs_entity_map_value.EOF Then
            ok_valor = False
        Else
            imgvalor = rs_entity_map_value!imgvalor
        End If
        rs_entity_map_value.Close
    End If
    
    If ok_valor Then
        StrSql = "SELECT codetosend FROM entity_value "
        StrSql = StrSql & " WHERE entity_value.entvalnro = " & imgvalor
        OpenRecordset StrSql, rs_entity_value
        If rs_entity_value.EOF Then
            ok_valor = False
        Else
            v_valor = rs_entity_value!codetosend
        End If
        rs_entity_value.Close
    End If
    
End Sub

Private Sub Depeve48(ByVal intnro As Integer, ByVal Empnro As Integer, ByVal ternro As Long, ByVal eventnro As Integer, ByVal topicnro As Integer, ByVal tfnro As Integer, ByRef v_tipovalor As Integer, ByRef v_valor As String, ByRef ok_valor As Boolean, ByRef msg_valor As String)
 Dim intfhasta As Date
 Dim intfdesde As Date
 Dim caunro As Integer
 Dim imgvalor As Integer
 
 Dim rs_interfaz As New ADODB.Recordset
 Dim rs_fases As New ADODB.Recordset
 Dim rs_entity_map_value As New ADODB.Recordset
 Dim rs_entity_value As New ADODB.Recordset
 
 ' v_tipovalor = 1 Caracter
 ' v_tipovalor = 2 Código interno de valor válido (bnp_entity_value.entvalnro)
 
    v_tipovalor = 1
    ok_valor = True
    msg_valor = ""
    
    ' Buscar el periodo de interfaz
    StrSql = "SELECT * FROM interfaz WHERE intnro = " & intnro
    OpenRecordset StrSql, rs_interfaz
    If rs_interfaz.EOF Then
        ok_valor = False
    Else
        intfhasta = rs_interfaz!intfhasta
        intfdesde = rs_interfaz!intfdesde
    End If
    rs_interfaz.Close
    
    If ok_valor Then
        ' Buscar baja dentro del período de la interfaz
        StrSql = "SELECT caunro FROM fases "
        StrSql = StrSql & " WHERE fases.empleado = " & ternro
        StrSql = StrSql & " AND bajfec >= " & ConvFecha(intfdesde) & " AND bajfec <= " & ConvFecha(intfhasta)
        StrSql = StrSql & " AND estado = 0 ORDER BY altfec DESC"
        OpenRecordset StrSql, rs_fases
        If rs_fases.EOF Then
            ok_valor = False
        Else
            caunro = rs_fases!caunro
        End If
        rs_fases.Close
    End If
    
    If ok_valor Then
        StrSql = "SELECT imgvalor FROM entity_map_value "
        StrSql = StrSql & " INNER JOIN entity_map ON entity_map_value.entmapnro = entity_map.entmapnro"
        StrSql = StrSql & " WHERE entity_map_value.domvalor = " & caunro & " AND entity_map.entmapnro = 14"
        OpenRecordset StrSql, rs_entity_map_value
        If rs_entity_map_value.EOF Then
            ok_valor = False
        Else
            imgvalor = rs_entity_map_value!imgvalor
        End If
        rs_entity_map_value.Close
    End If
    
    If ok_valor Then
        StrSql = "SELECT codetosend FROM entity_value "
        StrSql = StrSql & " WHERE entity_value.entvalnro = " & imgvalor
        OpenRecordset StrSql, rs_entity_value
        If rs_entity_value.EOF Then
            ok_valor = False
        Else
            v_valor = rs_entity_value!codetosend
        End If
        rs_entity_value.Close
    End If
    
End Sub

Private Sub Depeve49(ByVal intnro As Integer, ByVal Empnro As Integer, ByVal ternro As Long, ByVal eventnro As Integer, ByVal topicnro As Integer, ByVal tfnro As Integer, ByRef v_tipovalor As Integer, ByRef v_valor As String, ByRef ok_valor As Boolean, ByRef msg_valor As String)
 Dim rs_interfaz As New ADODB.Recordset
 
    v_tipovalor = 1
    ok_valor = True
    msg_valor = ""

    StrSql = "SELECT * FROM interfaz WHERE intnro = " & intnro
    OpenRecordset StrSql, rs_interfaz
    If rs_interfaz.EOF Then
        ok_valor = False
    Else
        Call formato(CStr(rs_interfaz!intfhasta), tfnro, v_tipovalor, v_valor)
    End If
    rs_interfaz.Close
    
End Sub

Private Sub Depeve50(ByVal intnro As Integer, ByVal Empnro As Integer, ByVal ternro As Long, ByVal eventnro As Integer, ByVal topicnro As Integer, ByVal tfnro As Integer, ByRef v_tipovalor As Integer, ByRef v_valor As String, ByRef ok_valor As Boolean, ByRef msg_valor As String)
 Dim intfhasta As Date
 Dim intfdesde As Date
 Dim inthhasta As String
 Dim inthdesde As String
 Dim caudnro As Integer
 Dim aud_campnro As Integer
 Dim acnro As Integer
 Dim salir As Boolean
 Dim salir_ant As Boolean
 Dim hora
 Dim aud_hor As Long
 Dim aud_fec As Date
 Dim aud_actual As String
 Dim audnro As Long
 
 Dim rs_interfaz As New ADODB.Recordset
 Dim rs_intauditoria As New ADODB.Recordset
 Dim rs_auditoria As New ADODB.Recordset
 Dim rs_auditoria_ant As New ADODB.Recordset
 Dim rs_Event_Topic_Field As New ADODB.Recordset
 
    v_tipovalor = 3
    ok_valor = True
    msg_valor = ""
    v_valor = ""

    StrSql = "SELECT * FROM interfaz WHERE intnro = " & intnro
    OpenRecordset StrSql, rs_interfaz
    If rs_interfaz.EOF Then
        ok_valor = False
    Else
        intfhasta = rs_interfaz!intfhasta
        intfdesde = rs_interfaz!intfdesde
        inthhasta = CStr(rs_interfaz!inthhasta)
        If Len(inthhasta) = 4 Then
            inthhasta = inthhasta & "00"
        End If
        inthdesde = CStr(rs_interfaz!inthdesde)
        If Len(inthdesde) = 4 Then
            inthdesde = inthdesde & "00"
        End If
    End If
    rs_interfaz.Close

    If ok_valor Then
        ok_valor = False
        
        ' Verificar si hubo actualización de este dato en el período de la interfaz
        ' y en caso afirmativo obtener el valor
        StrSql = "SELECT * FROM intauditoria "
        StrSql = StrSql & " WHERE tfnro = " & tfnro
        OpenRecordset StrSql, rs_intauditoria
        Do Until rs_intauditoria.EOF Or ok_valor
            caudnro = CInt(rs_intauditoria!caudnro)
            aud_campnro = CInt(rs_intauditoria!aud_campnro)
            acnro = CInt(rs_intauditoria!acnro)
            
            StrSql = " SELECT * FROM auditoria"
            StrSql = StrSql & " WHERE aud_ternro = " & ternro
            'StrSql = StrSql & " AND aud_emp = " & empnro
            StrSql = StrSql & " AND caudnro = " & caudnro
            StrSql = StrSql & " AND aud_campnro = " & aud_campnro
            StrSql = StrSql & " AND acnro = " & acnro
            StrSql = StrSql & " AND aud_fec >= " & ConvFecha(intfdesde)
            StrSql = StrSql & " AND aud_fec <= " & ConvFecha(intfhasta)
            StrSql = StrSql & " ORDER BY aud_fec DESC"
            OpenRecordset StrSql, rs_auditoria
            If Not rs_auditoria.EOF Then
                salir = False
                Do Until rs_auditoria.EOF Or salir
                    hora = Split(rs_auditoria!aud_hor, ":")
                    aud_hor = CLng(hora(0))
                    aud_hor = aud_hor * CLng(100) + CLng(hora(1))
                    aud_hor = aud_hor * CLng(100) + CLng(hora(2))
                    aud_fec = rs_auditoria!aud_fec
                    If (aud_fec >= intfdesde And aud_fec <= intfhasta) Or (aud_fec = intfdesde And aud_hor >= inthdesde) Or (aud_fec = intfhasta And aud_hor <= inthhasta) Then
                        salir = True
                        aud_actual = rs_auditoria!aud_actual
                        audnro = rs_auditoria!audnro
                    End If
                    
                    rs_auditoria.MoveNext
                Loop
                
                If salir Then
                    ' Verificar que el campo ya no halla tenido este valor en el mismo período,
                    ' si esto sucediera, el campo nunca cambio de valor
                    StrSql = " SELECT * FROM auditoria"
                    StrSql = StrSql & " WHERE aud_ternro = " & ternro
                    'StrSql = StrSql & " AND aud_emp = " & empnro
                    StrSql = StrSql & " AND caudnro = " & caudnro
                    StrSql = StrSql & " AND aud_campnro = " & aud_campnro
                    StrSql = StrSql & " AND acnro = " & acnro
                    StrSql = StrSql & " AND aud_fec >= " & ConvFecha(intfdesde)
                    StrSql = StrSql & " AND aud_fec <= " & ConvFecha(intfhasta)
                    StrSql = StrSql & " AND aud_actual = '" & aud_actual & "'"
                    StrSql = StrSql & " AND audnro <> " & audnro
                    StrSql = StrSql & " ORDER BY aud_fec DESC"
                    OpenRecordset StrSql, rs_auditoria_ant
                    salir_ant = False
                    Do Until rs_auditoria_ant.EOF Or salir_ant
                        hora = Split(rs_auditoria_ant!aud_hor, ":")
                        aud_hor = CLng(hora(0))
                        aud_hor = aud_hor * CLng(100) + CLng(hora(1))
                        aud_hor = aud_hor * CLng(100) + CLng(hora(2))
                        aud_fec = rs_auditoria!aud_fec
                        If Not ((aud_fec >= intfdesde And aud_fec <= intfhasta) Or (aud_fec = intfdesde And aud_hor >= inthdesde) Or (aud_fec = intfhasta And aud_hor <= inthhasta)) Then
                            salir_ant = True
                        End If
                        
                        rs_auditoria_ant.MoveNext
                    Loop
                    
                    If Not salir_ant Then
                        ok_valor = True
                    End If
                    
                    rs_auditoria_ant.Close
                    
                End If
                
            End If
            
            rs_auditoria.Close
            
            rs_intauditoria.MoveNext
        Loop
        
        rs_intauditoria.Close
        
    End If
 
    If ok_valor Then
        ' Buscar el programa que informa este dato en el evento de envío inicial (DTC_XXX)
        StrSql = "SELECT valueprg FROM event_topic_field "
        StrSql = StrSql & " WHERE event_topic_field.eventnro = 24 "     ' DTC_HIR
        StrSql = StrSql & " AND event_topic_field.tfnro = " & tfnro
        OpenRecordset StrSql, rs_Event_Topic_Field
        If Not rs_Event_Topic_Field.EOF Then
            Call obtener_valor(CStr(rs_Event_Topic_Field!valueprg), intnro, Empnro, ternro, eventnro, topicnro, tfnro, 0, "", "", "", v_tipovalor, v_valor, ok_valor, msg_valor)
        End If
        rs_Event_Topic_Field.Close
    End If
    
End Sub

Private Sub Depeve51(ByVal intnro As Integer, ByVal Empnro As Integer, ByVal ternro As Long, ByVal eventnro As Integer, ByVal topicnro As Integer, ByVal tfnro As Integer, ByRef v_tipovalor As Integer, ByRef v_valor As String, ByRef ok_valor As Boolean, ByRef msg_valor As String)

 ' FEX " Expatriation "     E " Expatriate "
 ' DET " Secondment "       D " Seconded "
 ' MAD " On loan "          M " On loan "
 
    v_tipovalor = 1
    ok_valor = True
    msg_valor = ""
    
    ' Buscar la clase del empleado y determinar la razon
    Call Depeve26(intnro, Empnro, ternro, eventnro, 2, 321, v_tipovalor, v_valor, ok_valor, msg_valor)
    
    Select Case v_valor
        Case "E": v_valor = "FEX"
        Case "D": v_valor = "DET"
        Case "M": v_valor = "MAD"
    End Select
        
    v_tipovalor = 2
    ok_valor = True
    msg_valor = ""
        
End Sub

Private Sub Depeve52(ByVal intnro As Integer, ByVal Empnro As Integer, ByVal ternro As Long, ByVal eventnro As Integer, ByVal topicnro As Integer, ByVal tfnro As Integer, ByRef v_tipovalor As Integer, ByRef v_valor As String, ByRef ok_valor As Boolean, ByRef msg_valor As String)
 Dim rs_interfaz As New ADODB.Recordset
 Dim rs_Event_Topic_Field As New ADODB.Recordset
 
    v_tipovalor = 3
    ok_valor = True
    msg_valor = ""
    v_valor = ""

    StrSql = "SELECT * FROM interfaz WHERE intnro = " & intnro
    OpenRecordset StrSql, rs_interfaz
    If rs_interfaz.EOF Then
        ok_valor = False
    End If
    rs_interfaz.Close
 
    If ok_valor Then
        ' Buscar el programa que informa este dato en el evento de envío inicial (DTC_XXX)
        StrSql = "SELECT valueprg FROM event_topic_field "
        StrSql = StrSql & " WHERE event_topic_field.eventnro = 24 "     ' DTC_HIR
        StrSql = StrSql & " AND event_topic_field.tfnro = " & tfnro
        OpenRecordset StrSql, rs_Event_Topic_Field
        If Not rs_Event_Topic_Field.EOF Then
            Call obtener_valor(CStr(rs_Event_Topic_Field!valueprg), intnro, Empnro, ternro, eventnro, topicnro, tfnro, 0, "", "", "", v_tipovalor, v_valor, ok_valor, msg_valor)
        End If
        rs_Event_Topic_Field.Close
    End If
    
End Sub


Private Sub regeve01(ByVal intnro As Integer, ByVal Empnro As Integer, ByVal ternro As Long, ByVal eventnro As Integer, ByVal topicnro As Integer, ByVal tfnro As Integer, ByVal nrofila As Integer, ByVal v As String, ByRef nuevo As Boolean)
 Dim rs_field_value As New ADODB.Recordset
 
    StrSql = "SELECT * FROM field_value "
    StrSql = StrSql & " WHERE field_value.tfnro = " & tfnro
    StrSql = StrSql & " AND field_value.ternro = " & ternro
    StrSql = StrSql & " AND field_value.intnro = " & intnro
    StrSql = StrSql & " AND field_value.eventnro = " & eventnro
    'StrSql = StrSql & " AND field_value.empnro = " & empnro
    OpenRecordset StrSql, rs_field_value
    If rs_field_value.EOF Then
        StrSql = "INSERT INTO field_value (tfnro,ternro,intnro,empnro,eventnro,topicnro,filanro,valor) "
        StrSql = StrSql & " VALUES (" & tfnro & "," & ternro & "," & intnro & "," & Empnro
        StrSql = StrSql & "," & eventnro & "," & topicnro & "," & nrofila & ",'" & Left(v, 100) & "')"
        
        nuevo = True
    Else
        StrSql = "UPDATE field_value SET valor = '" & Left(v, 100) & "'"
        StrSql = StrSql & " WHERE field_value.tfnro = " & tfnro
        StrSql = StrSql & " AND field_value.ternro = " & ternro
        StrSql = StrSql & " AND field_value.intnro = " & intnro
        StrSql = StrSql & " AND field_value.eventnro = " & eventnro
        'StrSql = StrSql & " AND field_value.empnro = " & empnro
    End If
    
    rs_field_value.Close
    
    objConn.Execute StrSql, , adExecuteNoRecords
    
End Sub

Private Sub regeve02(ByVal intnro As Integer, ByVal Empnro As Integer, ByVal ternro As Long, ByVal eventnro As Integer, ByVal topicnro As Integer)
 
    StrSql = "DELETE FROM field_value "
    StrSql = StrSql & " WHERE field_value.topicnro = " & topicnro
    StrSql = StrSql & " AND field_value.ternro = " & ternro
    StrSql = StrSql & " AND field_value.intnro = " & intnro
    'StrSql = StrSql & " AND field_value.empnro = " & empnro
    StrSql = StrSql & " AND field_value.eventnro = " & eventnro
    
    objConn.Execute StrSql, , adExecuteNoRecords
    
End Sub

Private Sub gendat01(ByVal intnro As Integer, ByVal Empnro As Integer, ByVal ternro As Long)
 Dim rs_empgenerado As New ADODB.Recordset
 
    StrSql = "SELECT * FROM empgenerado "
    StrSql = StrSql & " WHERE empgenerado.intnro = " & intnro
    'StrSql = StrSql & " AND empgenerado.empnro = " & empnro
    StrSql = StrSql & " AND empgenerado.ternro = " & ternro
    OpenRecordset StrSql, rs_empgenerado
    If rs_empgenerado.EOF Then
        StrSql = "INSERT INTO empgenerado (intnro, empnro, ternro) "
        StrSql = StrSql & " values (" & intnro & "," & Empnro & "," & ternro & ")"
        
        objConn.Execute StrSql, , adExecuteNoRecords
    End If
    
    rs_empgenerado.Close
    
End Sub

Private Sub obtener_valor(ByVal prog As String, ByVal intnro As Integer, ByVal Empnro As Integer, ByVal ternro As Long, ByVal eventnro As Integer, ByVal topicnro As Integer, ByVal tfnro As Integer, ByVal asoctype As Integer, ByVal eventcode As String, ByVal topicheader As String, ByVal tfname As String, ByRef v_tipovalor As Integer, ByRef v_valor As String, ByRef ok_valor As Boolean, ByRef msg_valor As String)
    Select Case UCase(prog)
        Case "INDEVE01.P":
            Call Indeve01(intnro, Empnro, ternro, eventnro, topicnro, tfnro, v_tipovalor, v_valor, ok_valor, msg_valor)
        Case "INDEVE02.P":
            Call Indeve02(intnro, Empnro, ternro, eventnro, topicnro, tfnro, v_tipovalor, v_valor, ok_valor, msg_valor)
        Case "INDEVE03.P":
            Call Indeve03(intnro, Empnro, ternro, eventnro, topicnro, tfnro, v_tipovalor, v_valor, ok_valor, msg_valor)
        Case "INDEVE04.P":
            Call Indeve04(intnro, Empnro, ternro, eventnro, topicnro, tfnro, v_tipovalor, v_valor, ok_valor, msg_valor)
        Case "INDEVE05.P":
            Call Indeve05(intnro, Empnro, ternro, eventnro, topicnro, tfnro, v_tipovalor, v_valor, ok_valor, msg_valor)
        Case "DEPEVE01.P":
            Call Depeve01(intnro, Empnro, ternro, eventnro, topicnro, tfnro, v_tipovalor, v_valor, ok_valor, msg_valor)
        Case "DEPEVE02.P":
            Call Depeve02(intnro, Empnro, ternro, eventnro, topicnro, tfnro, v_tipovalor, v_valor, ok_valor, msg_valor)
        Case "DEPEVE03.P":
            Call Depeve03(intnro, Empnro, ternro, eventnro, topicnro, tfnro, v_tipovalor, v_valor, ok_valor, msg_valor)
        Case "DEPEVE04.P":
            Call Depeve04(intnro, Empnro, ternro, eventnro, topicnro, tfnro, v_tipovalor, v_valor, ok_valor, msg_valor)
        Case "DEPEVE05.P":
            Call Depeve05(intnro, Empnro, ternro, eventnro, topicnro, tfnro, v_tipovalor, v_valor, ok_valor, msg_valor)
        Case "DEPEVE06.P":
            Call Depeve06(intnro, Empnro, ternro, eventnro, topicnro, tfnro, v_tipovalor, v_valor, ok_valor, msg_valor)
        Case "DEPEVE07.P":
            Call Depeve07(intnro, Empnro, ternro, eventnro, topicnro, tfnro, v_tipovalor, v_valor, ok_valor, msg_valor)
        Case "DEPEVE08.P":
            Call Depeve08(intnro, Empnro, ternro, eventnro, topicnro, tfnro, v_tipovalor, v_valor, ok_valor, msg_valor)
        Case "DEPEVE09.P":
            Call Depeve09(intnro, Empnro, ternro, eventnro, topicnro, tfnro, v_tipovalor, v_valor, ok_valor, msg_valor)
        Case "DEPEVE10.P":
            Call Depeve10(intnro, Empnro, ternro, eventnro, topicnro, tfnro, v_tipovalor, v_valor, ok_valor, msg_valor)
        Case "DEPEVE11.P":
            Call Depeve11(intnro, Empnro, ternro, eventnro, topicnro, tfnro, v_tipovalor, v_valor, ok_valor, msg_valor)
        Case "DEPEVE12.P":
            Call Depeve12(intnro, Empnro, ternro, eventnro, topicnro, tfnro, v_tipovalor, v_valor, ok_valor, msg_valor)
        Case "DEPEVE13.P":
            ' Hace lo mismo que el 05, por eso se llama a ese
            Call Depeve05(intnro, Empnro, ternro, eventnro, topicnro, tfnro, v_tipovalor, v_valor, ok_valor, msg_valor)
        Case "DEPEVE14.P":
            ' Hace lo mismo que el 06, por eso se llama a ese
            Call Depeve06(intnro, Empnro, ternro, eventnro, topicnro, tfnro, v_tipovalor, v_valor, ok_valor, msg_valor)
        Case "DEPEVE15.P":
            ' Hace lo mismo que el 07, por eso se llama a ese
            Call Depeve07(intnro, Empnro, ternro, eventnro, topicnro, tfnro, v_tipovalor, v_valor, ok_valor, msg_valor)
        Case "DEPEVE16.P":
            ' Hace lo mismo que el 08, por eso se llama a ese
            Call Depeve08(intnro, Empnro, ternro, eventnro, topicnro, tfnro, v_tipovalor, v_valor, ok_valor, msg_valor)
        Case "DEPEVE17.P":
            ' Hace lo mismo que el 09, por eso se llama a ese
            Call Depeve09(intnro, Empnro, ternro, eventnro, topicnro, tfnro, v_tipovalor, v_valor, ok_valor, msg_valor)
        Case "DEPEVE18.P":
            ' Hace lo mismo que el 10, por eso se llama a ese
            Call Depeve10(intnro, Empnro, ternro, eventnro, topicnro, tfnro, v_tipovalor, v_valor, ok_valor, msg_valor)
        Case "DEPEVE19.P":
            ' Hace lo mismo que el 11, por eso se llama a ese
            Call Depeve11(intnro, Empnro, ternro, eventnro, topicnro, tfnro, v_tipovalor, v_valor, ok_valor, msg_valor)
        Case "DEPEVE20.P":
            ' Hace lo mismo que el 12, por eso se llama a ese
            Call Depeve12(intnro, Empnro, ternro, eventnro, topicnro, tfnro, v_tipovalor, v_valor, ok_valor, msg_valor)
        Case "DEPEVE21.P":
            Call Depeve21(intnro, Empnro, ternro, eventnro, topicnro, tfnro, v_tipovalor, v_valor, ok_valor, msg_valor)
        Case "DEPEVE22.P":
            Call Depeve22(intnro, Empnro, ternro, eventnro, topicnro, tfnro, v_tipovalor, v_valor, ok_valor, msg_valor)
        Case "DEPEVE23.P":
            Call Depeve23(intnro, Empnro, ternro, eventnro, topicnro, tfnro, v_tipovalor, v_valor, ok_valor, msg_valor)
        Case "DEPEVE24.P":
            Call Depeve24(intnro, Empnro, ternro, eventnro, topicnro, tfnro, v_tipovalor, v_valor, ok_valor, msg_valor)
        Case "DEPEVE25.P":
            Call Depeve25(intnro, Empnro, ternro, eventnro, topicnro, tfnro, v_tipovalor, v_valor, ok_valor, msg_valor)
        Case "DEPEVE26.P":
            Call Depeve26(intnro, Empnro, ternro, eventnro, topicnro, tfnro, v_tipovalor, v_valor, ok_valor, msg_valor)
        Case "DEPEVE27.P":
            Call Depeve27(intnro, Empnro, ternro, eventnro, topicnro, tfnro, v_tipovalor, v_valor, ok_valor, msg_valor)
        Case "DEPEVE28.P":
            Call Depeve28(intnro, Empnro, ternro, eventnro, topicnro, tfnro, v_tipovalor, v_valor, ok_valor, msg_valor)
        Case "DEPEVE29.P":
            Call Depeve29(intnro, Empnro, ternro, eventnro, topicnro, tfnro, v_tipovalor, v_valor, ok_valor, msg_valor)
        Case "DEPEVE30.P":
            Call Depeve30(intnro, Empnro, ternro, eventnro, topicnro, tfnro, v_tipovalor, v_valor, ok_valor, msg_valor)
        Case "DEPEVE31.P":
            Call Depeve31(intnro, Empnro, ternro, eventnro, topicnro, tfnro, v_tipovalor, v_valor, ok_valor, msg_valor)
        Case "DEPEVE32.P":
            Call Depeve32(intnro, Empnro, ternro, eventnro, topicnro, tfnro, v_tipovalor, v_valor, ok_valor, msg_valor)
        Case "DEPEVE33.P":
            Call Depeve33(intnro, Empnro, ternro, eventnro, topicnro, tfnro, v_tipovalor, v_valor, ok_valor, msg_valor)
        Case "DEPEVE34.P":
            Call Depeve34(intnro, Empnro, ternro, eventnro, topicnro, tfnro, v_tipovalor, v_valor, ok_valor, msg_valor)
        Case "DEPEVE35.P":
            Call Depeve35(intnro, Empnro, ternro, eventnro, topicnro, tfnro, v_tipovalor, v_valor, ok_valor, msg_valor)
        Case "DEPEVE36.P":
            Call Depeve36(intnro, Empnro, ternro, eventnro, topicnro, tfnro, v_tipovalor, v_valor, ok_valor, msg_valor)
        Case "DEPEVE41.P":
            Call Depeve41(intnro, Empnro, ternro, eventnro, topicnro, tfnro, v_tipovalor, v_valor, ok_valor, msg_valor)
        Case "DEPEVE42.P":
            Call Depeve42(intnro, Empnro, ternro, eventnro, topicnro, tfnro, v_tipovalor, v_valor, ok_valor, msg_valor)
        Case "DEPEVE43.P":
            Call Depeve43(intnro, Empnro, ternro, eventnro, topicnro, tfnro, v_tipovalor, v_valor, ok_valor, msg_valor)
        Case "DEPEVE44.P":
            Call Depeve44(intnro, Empnro, ternro, eventnro, topicnro, tfnro, v_tipovalor, v_valor, ok_valor, msg_valor)
        Case "DEPEVE45.P":
            Call Depeve45(intnro, Empnro, ternro, eventnro, topicnro, tfnro, v_tipovalor, v_valor, ok_valor, msg_valor)
        Case "DEPEVE46.P":
            Call Depeve46(intnro, Empnro, ternro, eventnro, topicnro, tfnro, v_tipovalor, v_valor, ok_valor, msg_valor)
        Case "DEPEVE47.P":
            Call Depeve47(intnro, Empnro, ternro, eventnro, topicnro, tfnro, v_tipovalor, v_valor, ok_valor, msg_valor)
        Case "DEPEVE48.P":
            Call Depeve48(intnro, Empnro, ternro, eventnro, topicnro, tfnro, v_tipovalor, v_valor, ok_valor, msg_valor)
        Case "DEPEVE49.P":
            Call Depeve49(intnro, Empnro, ternro, eventnro, topicnro, tfnro, v_tipovalor, v_valor, ok_valor, msg_valor)
        Case "DEPEVE50.P":
            Call Depeve50(intnro, Empnro, ternro, eventnro, topicnro, tfnro, v_tipovalor, v_valor, ok_valor, msg_valor)
        Case "DEPEVE51.P":
            Call Depeve51(intnro, Empnro, ternro, eventnro, topicnro, tfnro, v_tipovalor, v_valor, ok_valor, msg_valor)
        Case "DEPEVE52.P":
            Call Depeve52(intnro, Empnro, ternro, eventnro, topicnro, tfnro, v_tipovalor, v_valor, ok_valor, msg_valor)
        Case Else
            Flog.writeline Espacios(Tabulador * 5) & "No esta definido el programa de busqueda (" & prog & " )."
            ' Error, Warning
            If asoctype = 1 Then
                Call registro_log(intnro, 1, ternro, 2, eventcode & ":" & topicheader + ":" + tfname & "-Programa Valor no encontrado: " & prog)
            Else
                Call registro_log(intnro, 1, ternro, 1, eventcode & ":" & topicheader + ":" + tfname & "-Programa Valor no encontrado: " & prog)
            End If
            ' rs_Event_Topic_Field!asoctype, rs_event!eventcode, rs_Event_Topic!topicheader , rs_topic_field!tfname
    End Select
End Sub
Private Sub formato(ByVal valor_in As String, ByVal tfnro As Integer, ByRef v_tipovalor As Integer, ByRef v_valor As String)
 Dim rs_topic_field As New ADODB.Recordset
 
    StrSql = "SELECT data_type.dtconvprogram, topic_field.tftotlength, topic_field.tfdecimals FROM topic_field "
    StrSql = StrSql & " INNER JOIN data_type ON data_type.dtnro = topic_field.dtnro "
    StrSql = StrSql & " WHERE topic_field.tfnro = " & tfnro
    
    OpenRecordset StrSql, rs_topic_field
    
    Select Case UCase(rs_topic_field!dtconvprogram)
        Case "DTFDATE.P":
            Call dtfdate(valor_in, rs_topic_field!tftotlength, rs_topic_field!tfdecimals, v_valor)
        Case "DTCCHAR.P":
            Call dtcchar(valor_in, rs_topic_field!tftotlength, rs_topic_field!tfdecimals, v_valor)
        Case "DTFNUMBER.P":
            Call dtfnumber(CDbl(valor_in), rs_topic_field!tftotlength, rs_topic_field!tfdecimals, v_valor)
    End Select
    
    rs_topic_field.Close
    
    v_tipovalor = 2
    
End Sub

Private Sub dtfdate(ByVal cin As String, ByVal l As Integer, ByVal d As Integer, ByRef cout As String)
    cout = CStr(Format(cin, "yyyy-mm-dd"))
End Sub

Private Sub dtcchar(ByVal cin As String, ByVal l As Integer, ByVal d As Integer, ByRef cout As String)
 Dim I As Integer
 
    cout = cin
    ' Eliminar caracteres no imprimibles
    For I = 0 To 31
        cout = Replace(cout, Chr(I), "")
    Next
    ' Espacio      32
    ' "!"          33
    For I = 34 To 36
        cout = Replace(cout, Chr(I), "")
    Next
    ' "%"          37
    ' "&"          38
    cout = Replace(cout, Chr(39), "")
    ' "("          40
    ' ")"          41
    ' "*"          42
    ' "+"          43
    ' ","          44
    ' "-"          45
    ' "."          46
    ' "/"          47
    ' "0" - "9"    48-57
    ' ":"          58
    ' ";"          59
    ' "<"          60
    cout = Replace(cout, Chr(61), "")
    ' ">"          62
    ' "?"          63
    cout = Replace(cout, Chr(64), "")
    '* "A" - "Z"    65-90
    For I = 91 To 96
        cout = Replace(cout, Chr(I), "")
    Next
    ' "a" - "z"    97-122
    For I = 123 To 191
        cout = Replace(cout, Chr(I), "")
    Next
    For I = 192 To 197
        cout = Replace(cout, Chr(I), "A")
    Next
    cout = Replace(cout, Chr(198), "")
    cout = Replace(cout, Chr(199), "")
    For I = 200 To 203
        cout = Replace(cout, Chr(I), "E")
    Next
    For I = 204 To 207
        cout = Replace(cout, Chr(I), "I")
    Next
    cout = Replace(cout, Chr(208), "")
    cout = Replace(cout, Chr(209), "N")
    For I = 210 To 214
        cout = Replace(cout, Chr(I), "O")
    Next
    For I = 215 To 216
        cout = Replace(cout, Chr(I), "")
    Next
    For I = 217 To 220
        cout = Replace(cout, Chr(I), "U")
    Next
    cout = Replace(cout, Chr(221), "Y")
    cout = Replace(cout, Chr(222), "")
    cout = Replace(cout, Chr(223), "")
    For I = 224 To 229
        cout = Replace(cout, Chr(I), "a")
    Next
    cout = Replace(cout, Chr(230), "")
    cout = Replace(cout, Chr(231), "")
    For I = 232 To 235
        cout = Replace(cout, Chr(I), "e")
    Next
    For I = 236 To 239
        cout = Replace(cout, Chr(I), "i")
    Next
    cout = Replace(cout, Chr(240), "")
    cout = Replace(cout, Chr(241), "n")
    For I = 242 To 246
        cout = Replace(cout, Chr(I), "o")
    Next
    For I = 247 To 248
        cout = Replace(cout, Chr(I), "")
    Next
    For I = 249 To 252
        cout = Replace(cout, Chr(I), "u")
    Next
    cout = Replace(cout, Chr(253), "y")
    cout = Replace(cout, Chr(254), "")
    cout = Replace(cout, Chr(255), "y")
End Sub

Private Sub dtfnumber(ByVal n As Double, ByVal l As Integer, ByVal d As Integer, ByRef nout As String)
    If n = 0 Or EsNulo(n) Then
        nout = "0"
    Else
        n = Round(n, d)
        nout = CStr(n)
        'nout = Replace(nout, ".", ",")
        'If InStr(nout, ".") = 0 Then
        '    nout = nout & ".0"
        'End If
    End If
    
'ASSIGN n = Round(n, d)
'       nout = STRING(n)
'       nout = Replace(nout, ".", ",")
'       pdec = IF INDEX(nout, ",") = 0 THEN "" ELSE TRIM(ENTRY(2, nout))
'       pent = Trim(ENTRY(1, nout))
'       pent = IF pent = "" THEN "0" ELSE pent
'       nout = pent + (IF pdec = "" THEN "" ELSE ("." + pdec)).
End Sub

Private Sub event_excluido(ByVal intnro As Integer, ByVal Empnro As Integer, ByVal ternro As Long, ByVal eventnro As Integer, ByRef excluido As Boolean)
 Dim rs As New ADODB.Recordset
 
    StrSql = "SELECT * FROM event_rule "
    StrSql = StrSql & " INNER JOIN field_value ON event_rule.eventnro2 = field_value.eventnro "
    StrSql = StrSql & " WHERE event_rule.eventnro1 = " & eventnro
    StrSql = StrSql & " AND field_value.intnro = " & intnro
    StrSql = StrSql & " AND field_value.ternro = " & ternro
    'StrSql = StrSql & " AND field_value.empnro = " & empnro
    OpenRecordset StrSql, rs
    If rs.EOF Then
        excluido = False
    Else
        excluido = True
    End If
    rs.Close
    
    Set rs = Nothing
End Sub

Private Sub registro_log(ByVal intnro As Integer, ByVal Empnro As Integer, ByVal ternro As Long, ByVal tipo As Integer, ByVal msg As String)
'--------------------------------------------------------------------------------
' Descripción: devuelve un string con el valor por defecto definido para ese campo
' Autor      : Fernando Favre
' Fecha      : 13/05/2005
'   tipo    1 - Warning
'           2 - Error
'-------------------------------------------------------------------------------
    
    StrSql = "INSERT INTO intemplog (empnro,ternro,intnro,logentrydate,logentrytime,logentrytype,logentrymsg) "
    StrSql = StrSql & " VALUES (" & Empnro & "," & ternro & "," & intnro & ","
    StrSql = StrSql & ConvFecha(Date) & ",'" & Format(Time, "HH:mm:") & "',"
    StrSql = StrSql & tipo & ",'" & Left(msg, 100) & "')"
    
    objConn.Execute StrSql, , adExecuteNoRecords
    
End Sub
