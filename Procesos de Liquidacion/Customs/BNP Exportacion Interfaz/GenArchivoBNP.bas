Attribute VB_Name = "GenArchivoBNP"
Option Explicit

Global Const Version = "1.01"
Global Const FechaModificacion = "31/03/2006"
Global Const UltimaModificacion = " " 'Version Inicial

'-----------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------

Global fs, f
Global Flog
Global NroProceso As Long

Global Path As String
Global HuboErrores As Boolean

Private Sub Main()
' ---------------------------------------------------------------------------------------------
' Descripcion : Procedimiento inicial exportacion
' Autor       : Fernando Favre
' Fecha       : 31/03/2006
' Ultima Mod  :
' Descripcion :
' ---------------------------------------------------------------------------------------------
Dim strCmdLine As String
Dim Nombre_Arch As String

Dim rs As New ADODB.Recordset
Dim PID As String
Dim Parametros As String
Dim ArrParametros

Dim intnro As Long
Dim empresa As Long


    strCmdLine = Command()
    ArrParametros = Split(strCmdLine, " ", -1)
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
    
    ' carga las configuraciones basicas, formato de fecha, string de conexion,
    ' tipo de BD y ubicacion del archivo de log
    Call CargarConfiguracionesBasicas

    On Error Resume Next
    'Abro la conexion
    OpenConnection strconexion, objConn
    If Err.Number <> 0 Then
        Flog.writeline "Problemas en la conexion"
        Exit Sub
    End If
    OpenConnection strconexion, objconnProgreso
    If Err.Number <> 0 Then
        Flog.writeline "Problemas en la conexion"
        Exit Sub
    End If
    On Error GoTo 0

    On Error GoTo ME_Main
    
    TiempoInicialProceso = GetTickCount
    HuboErrores = False
    
    Nombre_Arch = PathFLog & "ExpInterfaz" & "-" & NroProceso & ".log"
    
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
    Flog.writeline "PID = " & PID
   
    'Cambio el estado del proceso a Procesando
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcestado = 'Procesando', bprcpid = " & PID & ", bprcprogreso = 0 WHERE bpronro = " & NroProceso
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    Flog.writeline Espacios(Tabulador * 0) & "Obtengo los datos del proceso"
    
    TiempoAcumulado = GetTickCount
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE bpronro = " & NroProceso
    StrSql = StrSql & " AND btprcnro = 127"
    OpenRecordset StrSql, rs
    If Not rs.EOF Then
       'Obtengo los parametros del proceso
       Parametros = rs!bprcparam
       ArrParametros = Split(Parametros, "@")
       
       intnro = CLng(ArrParametros(0))
       empresa = CLng(ArrParametros(1))
      
       Call Generar_Archivo(intnro, empresa)
    Else
        Flog.writeline Espacios(Tabulador * 0) & "No se encontraron los datos del proceso nro " & NroProceso & " de tipo 127."
    End If
    
    'Actualizo el estado de la interfaz
    If Not HuboErrores Then
        StrSql = "UPDATE interfaz SET "
        StrSql = StrSql & " intfultgenarch = " & ConvFecha(Now)
        StrSql = StrSql & " ,inthultgenarch = " & Format(Time, "HHMMSS")
        StrSql = StrSql & " WHERE intnro = " & intnro
        objConn.Execute StrSql, , adExecuteNoRecords
    End If
    
    'Actualizo el estado del proceso
    If Not HuboErrores Then
       StrSql = "UPDATE batch_proceso SET  bprcprogreso =100, bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado' WHERE bpronro = " & NroProceso
       Flog.writeline Espacios(Tabulador * 0) & "Proceso Finalizado Correctamente"
    Else
       StrSql = "UPDATE batch_proceso SET  bprcprogreso =100, bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Incompleto' WHERE bpronro = " & NroProceso
       Flog.writeline Espacios(Tabulador * 0) & "Proceso Finalizado Incompleto"
    End If
    objconnProgreso.Execute StrSql, , adExecuteNoRecords

    TiempoFinalProceso = GetTickCount
    Flog.writeline Espacios(Tabulador * 0) & "=================================================="
    Flog.writeline Espacios(Tabulador * 0) & "Tiempo del proceso (milisegundos): " & (TiempoFinalProceso - TiempoInicialProceso)
    Flog.writeline Espacios(Tabulador * 0) & "=================================================="
    Flog.Close
    If rs.State = adStateOpen Then rs.Close
    Set rs = Nothing
    objconnProgreso.Close
    objConn.Close
Exit Sub
    
ME_Main:
    HuboErrores = True
    Flog.writeline "Error: " & Err.Description
    Flog.writeline "Ultima SQL ejecutada: " & StrSql
End Sub

Private Sub Generar_Archivo(ByVal intnro As Long, ByVal empresa As Long)
' ---------------------------------------------------------------------------------------------
' Descripcion : Procedimiento que genera la exportacion
' Autor       : Fernando Favre
' Fecha       : 31/03/2006
' Ultima Mod  :
' Descripcion :
' ---------------------------------------------------------------------------------------------
'Dim objRs As New ADODB.Recordset
Dim Nombre_Arch As String
Dim NroModelo As Long
Dim Directorio As String
Dim ArchExp
Dim Carpeta
Dim fs1

Dim Sep As String
Dim SepDec As String

Dim cantRegistros As Long

Dim lista_emp As String
Dim cont As Integer
Dim Valor(6) As String
Dim separador As String
Dim intnomarch As String
Dim intfhasta As String
Dim v_linea As Integer
Dim topicorder_ant As Integer

Dim sysid As String
Dim linea As String
Dim v As String

Dim rs_field_value As New ADODB.Recordset
Dim rs_topic As New ADODB.Recordset
Dim rs_topic_field As New ADODB.Recordset
Dim rs_field_value2 As New ADODB.Recordset
Dim rs_intconfgen As New ADODB.Recordset
Dim rs_interfaz As New ADODB.Recordset
Dim rs_tt_topicline As New ADODB.Recordset
Dim rs_tt_line As New ADODB.Recordset

    On Error GoTo ME_Local
      
    'NroModelo = 273
    'Directorio de exportacion
    StrSql = "SELECT sis_dirsalidas FROM sistema WHERE sisnro = 1 "
    If rs.State = adStateOpen Then rs.Close
    OpenRecordset StrSql, rs
    If Not rs.EOF Then
        Directorio = Trim(rs!sis_dirsalidas)
    End If
    rs.Close
    
    'StrSql = "SELECT * FROM modelo WHERE modnro = " & NroModelo
    'OpenRecordset StrSql, rs_Modelo
    'If Not rs_Modelo.EOF Then
    '    If Not IsNull(rs_Modelo!modarchdefault) Then
    '        Directorio = Directorio & Trim(rs_Modelo!modarchdefault)
    '    Else
    '        Flog.writeline Espacios(Tabulador * 1) & "El modelo no tiene configurada la carpeta desteino. El archivo será generado en el directorio default"
    '    End If
    'Else
    '    Flog.writeline Espacios(Tabulador * 1) & "No se encontró el modelo " & NroModelo & ". El archivo será generado en el directorio default"
    'End If
             
    'Obtengo los datos del separador
    'Sep = rs_Modelo!modseparador
    'SepDec = rs_Modelo!modsepdec
    'Flog.writeline Espacios(Tabulador * 0) & "Separador seleccionado: " & Sep
     
    
    ' Borro los datos de la tabla temporal
    StrSql = "DELETE FROM tt_line"
    objConn.Execute StrSql, , adExecuteNoRecords
    
    
    ' Borro los datos de la tabla temporal
    StrSql = "DELETE FROM tt_topicline"
    objConn.Execute StrSql, , adExecuteNoRecords
    
    
    ' Busca los datos de la interfaz
    StrSql = "SELECT * FROM interfaz "
    StrSql = StrSql & " WHERE intnro = " & intnro
    OpenRecordset StrSql, rs_interfaz
    If Not rs_interfaz.EOF Then
        intfhasta = CStr(Format(rs_interfaz!intfhasta, "yyyy-mm-dd"))
        intnomarch = rs_interfaz!intnomarch
    Else
        Flog.writeline Espacios(Tabulador * 1) & "No se encontraron los datos de la interfaz."
        GoTo Fin
    End If
    rs_interfaz.Close
    
    
    Nombre_Arch = Directorio & "\Interfaz\" & intnomarch
    Flog.writeline Espacios(Tabulador * 1) & "Se crea el archivo: " & Nombre_Arch
    Set fs = CreateObject("Scripting.FileSystemObject")
    'desactivo el manejador de errores
    On Error Resume Next
    
    Set Carpeta = fs.getFolder(Directorio)
    If Err.Number <> 0 Then
        Flog.writeline Espacios(Tabulador * 1) & "La carpeta Destino no existe. Se creará."
        Err.Number = 0
        Set Carpeta = fs.CreateFolder(Directorio)
        
        If Err.Number <> 0 Then
            Flog.writeline Espacios(Tabulador * 1) & "No se puede crear el directorio. Verifique los derechos de acceso o puede crearlo."
            GoTo Fin
        End If
    End If
    
    Set Carpeta = fs.getFolder(Directorio & "\Interfaz")
    If Err.Number <> 0 Then
        Flog.writeline Espacios(Tabulador * 1) & "El directorio " & Directorio & "\Interfaz no existe. Se creará."
        Err.Number = 0
        Set Carpeta = fs.CreateFolder(Directorio & "\Interfaz")
        
        If Err.Number <> 0 Then
            Flog.writeline Espacios(Tabulador * 1) & "No se puede crear el directorio. Verifique los derechos de acceso o puede crearlo."
            GoTo Fin
        End If
    End If

    Set ArchExp = fs.CreateTextFile(Nombre_Arch, True)
    
    If Err.Number <> 0 Then
        Flog.writeline Espacios(Tabulador * 1) & "No se puede crear el archivo " & Nombre_Arch & " en el directorio " & Directorio & "\Interfaz "
        GoTo Fin
    End If

    On Error GoTo ME_Local
    
    ' Ordenar los Datos
    StrSql = "SELECT * FROM field_value "
    StrSql = StrSql & " WHERE field_value.intnro = " & intnro
    'If empresa <> "-1" Then
    '    StrSql = StrSql & " AND field_value.empnro = " & empresa
    'End If
    StrSql = StrSql & " ORDER BY field_value.topicnro, field_value.eventnro, field_value.filanro, field_value.empnro, field_value.ternro"
    OpenRecordset StrSql, rs_field_value
    
    'seteo de las variables de progreso
    Progreso = 0
    If Not rs_field_value.EOF Then
        cantRegistros = rs_field_value.RecordCount
        If cantRegistros = 0 Then
           cantRegistros = 1
           Flog.writeline Espacios(Tabulador * 1) & "No se encontraron datos en la interfaz para Exportar. SQL: " & StrSql
        End If
    Else
        cantRegistros = 1
        Flog.writeline Espacios(Tabulador * 1) & "No se encontraron datos en la interfaz para Exportar. SQL: " & StrSql
    End If
    IncPorc = (49 / cantRegistros)
    
    lista_emp = ","
    Do Until rs_field_value.EOF
        ' Obtener los campos de ordenamiento para el tópico
        If InStr(1, lista_emp, "," & rs_field_value!topicnro & "@" & rs_field_value!eventnro & "@" & rs_field_value!filanro & "@" & rs_field_value!ternro & ",") = 0 Then
            cont = 1
            
            lista_emp = lista_emp & rs_field_value!topicnro & "@" & rs_field_value!eventnro & "@" & rs_field_value!filanro & "@" & rs_field_value!ternro & ","
            StrSql = "SELECT * FROM topic_field "
            StrSql = StrSql & " WHERE topic_field.topicnro = " & rs_field_value!topicnro
            StrSql = StrSql & " AND topic_field.ordena = -1 "
            StrSql = StrSql & " ORDER BY topic_field.tforden "
            OpenRecordset StrSql, rs_topic_field
            
            Do Until rs_topic_field.EOF
                StrSql = "SELECT * FROM field_value "
                StrSql = StrSql & " WHERE field_value.tfnro = " & rs_topic_field!tfnro
                StrSql = StrSql & " AND field_value.intnro = " & rs_field_value!intnro
                StrSql = StrSql & " AND field_value.ternro = " & rs_field_value!ternro
                'StrSql = StrSql & " AND field_value.empnro = " & rs_field_value!empnro
                StrSql = StrSql & " AND field_value.eventnro = " & rs_field_value!eventnro
                StrSql = StrSql & " AND field_value.filanro = " & rs_field_value!filanro
                OpenRecordset StrSql, rs_field_value2
                
                If Not rs_field_value2.EOF Then
                    Valor(cont) = rs_field_value2!Valor
                Else
                    Valor(cont) = ""
                End If
                
                rs_field_value2.Close
                
                rs_topic_field.MoveNext
                
                cont = cont + 1
            Loop
            
            rs_topic_field.Close
            
            StrSql = "INSERT INTO tt_topicline (intnro,ternro,empnro,tfnro,topicnro,eventnro,"
            StrSql = StrSql & "filanro,ordfield1,ordfield2,ordfield3,ordfield4,ordfield5,ordfield6) "
            StrSql = StrSql & " VALUES (" & rs_field_value!intnro
            StrSql = StrSql & "," & rs_field_value!ternro
            StrSql = StrSql & "," & rs_field_value!Empnro
            StrSql = StrSql & "," & rs_field_value!tfnro
            StrSql = StrSql & "," & rs_field_value!topicnro
            StrSql = StrSql & "," & rs_field_value!eventnro
            StrSql = StrSql & "," & rs_field_value!filanro
            StrSql = StrSql & ",'" & Valor(1)
            StrSql = StrSql & "','" & Valor(2)
            StrSql = StrSql & "','" & Valor(3)
            StrSql = StrSql & "','" & Valor(4)
            StrSql = StrSql & "','" & Valor(5)
            StrSql = StrSql & "','" & Valor(6) & "')"
            objConn.Execute StrSql, , adExecuteNoRecords
            
        End If
        rs_field_value.MoveNext
    
        'Actualizo el estado del proceso
        TiempoAcumulado = GetTickCount
        cantRegistros = cantRegistros - 1
        Progreso = Progreso + IncPorc
        
        StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso
        StrSql = StrSql & ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'"
        StrSql = StrSql & ", bprcempleados ='" & CStr(cantRegistros) & "' WHERE bpronro = " & NroProceso
        objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    Loop
    rs_field_value.Close
    
    
    ' Completar tópicos no existentes
    StrSql = "SELECT topicnro FROM topic "
    StrSql = StrSql & " WHERE topicnro NOT IN (SELECT topicnro FROM tt_topicline) "
    StrSql = StrSql & " ORDER BY topic.topicorder"
    OpenRecordset StrSql, rs_topic
    Do Until rs_topic.EOF
        StrSql = "INSERT INTO tt_topicline (tfnro,topicnro) "
        StrSql = StrSql & " VALUES (-1, " & rs_topic!topicnro & " )"
        objConn.Execute StrSql, , adExecuteNoRecords
        rs_topic.MoveNext
    Loop
    rs_topic.Close
    
    
    ' Busca la configuracion general del sistema para obtener el separador
    StrSql = "SELECT * FROM intconfgen "
    OpenRecordset StrSql, rs_intconfgen
    If Not rs_intconfgen.EOF Then
        separador = rs_intconfgen!sepstr
        sysid = rs_intconfgen!sysid
    End If
    rs_intconfgen.Close
    
    v_linea = 1
    
    linea = Trim(sysid) & separador & intfhasta
    
    Call NewLine(linea, v_linea)
    
    
    ' Para cada tema por persona
    StrSql = "SELECT topic.topicnro, topic.topicheader, tt_topicline.tfnro, tt_topicline.intnro, tt_topicline.ternro, "
    StrSql = StrSql & " tt_topicline.empnro, tt_topicline.eventnro, tt_topicline.filanro, tt_topicline.ordfield1, "
    StrSql = StrSql & " tt_topicline.ordfield2, tt_topicline.ordfield3, tt_topicline.ordfield4, tt_topicline.ordfield5, "
    StrSql = StrSql & " tt_topicline.ordfield6, topic.topicorder FROM tt_topicline "
    StrSql = StrSql & " INNER JOIN topic ON tt_topicline.topicnro = topic.topicnro "
    StrSql = StrSql & " ORDER BY topic.topicorder, tt_topicline.ordfield1, tt_topicline.ordfield2, "
    StrSql = StrSql & " tt_topicline.ordfield3, tt_topicline.ordfield4, tt_topicline.ordfield5, tt_topicline.ordfield6 "
    OpenRecordset StrSql, rs_tt_topicline
    
    'seteo de las variables de progreso
    Progreso = 49
    If Not rs_tt_topicline.EOF Then
        cantRegistros = rs_tt_topicline.RecordCount
        If cantRegistros = 0 Then
           cantRegistros = 1
           Flog.writeline Espacios(Tabulador * 1) & "No se encontraron datos en tt_topicline. SQL: " & StrSql
        End If
    Else
        cantRegistros = 1
        Flog.writeline Espacios(Tabulador * 1) & "No se encontraron datos en tt_topicline. SQL: " & StrSql
    End If
    IncPorc = (49 / cantRegistros)
    
    topicorder_ant = 0
    Do Until rs_tt_topicline.EOF
        If (CInt(topicorder_ant) <> CInt(rs_tt_topicline!topicorder)) Then
            ' Generar etiqueta de topico
            topicorder_ant = rs_tt_topicline!topicorder
            
            Call NewLine(CStr(rs_tt_topicline!topicheader), v_linea)
        End If
        
        ' Verificar si es una linea de tópico sin datos
        If CInt(rs_tt_topicline!tfnro) <> CInt(-1) Then
            linea = ""
            ' Recorro los campos de la topic - line
            StrSql = "SELECT * FROM topic_field "
            StrSql = StrSql & " WHERE topicnro = " & rs_tt_topicline!topicnro
            StrSql = StrSql & " ORDER BY tforden"
            OpenRecordset StrSql, rs_topic_field
            
            Do Until rs_topic_field.EOF
                StrSql = "SELECT field_value.valor FROM field_value "
                StrSql = StrSql & " WHERE tfnro = " & rs_topic_field!tfnro
                StrSql = StrSql & " AND intnro = " & rs_tt_topicline!intnro
                StrSql = StrSql & " AND ternro = " & rs_tt_topicline!ternro
                'StrSql = StrSql & " AND empnro = " & rs_tt_topicline!empnro
                StrSql = StrSql & " AND eventnro = " & rs_tt_topicline!eventnro
                StrSql = StrSql & " AND filanro = " & rs_tt_topicline!filanro
                OpenRecordset StrSql, rs_field_value
                
                If rs_field_value.EOF Then
                    v = ""
                Else
                    v = rs_field_value!Valor
                    
                    StrSql = "UPDATE field_value SET archivado = -1 "
                    StrSql = StrSql & " WHERE tfnro = " & rs_topic_field!tfnro
                    StrSql = StrSql & " AND intnro = " & rs_tt_topicline!intnro
                    StrSql = StrSql & " AND ternro = " & rs_tt_topicline!ternro
                    'StrSql = StrSql & " AND empnro = " & rs_tt_topicline!empnro
                    StrSql = StrSql & " AND eventnro = " & rs_tt_topicline!eventnro
                    StrSql = StrSql & " AND filanro = " & rs_tt_topicline!filanro
                    objConn.Execute StrSql, , adExecuteNoRecords
                End If
                
                linea = linea & Trim(v) & separador
                
                rs_topic_field.MoveNext
            Loop
            
            linea = Mid(linea, 1, Len(linea) - 1)
            
            Call NewLine(linea, v_linea)
            
            rs_topic_field.Close
            
        End If
        
        rs_tt_topicline.MoveNext
        
        'Actualizo el estado del proceso
        TiempoAcumulado = GetTickCount
        cantRegistros = cantRegistros - 1
        Progreso = Progreso + IncPorc
        
        StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso
        StrSql = StrSql & ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'"
        StrSql = StrSql & ", bprcempleados ='" & CStr(cantRegistros) & "' WHERE bpronro = " & NroProceso
        objconnProgreso.Execute StrSql, , adExecuteNoRecords
        
    Loop
    
    rs_tt_topicline.Close
    
    
    StrSql = "SELECT * FROM tt_line ORDER BY line_number "
    OpenRecordset StrSql, rs_tt_line
    
    Do Until rs_tt_line.EOF
        ' Escribo en el archivo
        Call imprimirTexto(rs_tt_line!line_text, ArchExp, Len(rs_tt_line!line_text), True)
        
        rs_tt_line.MoveNext
    
        If Not rs_tt_line.EOF Then
            ArchExp.writeline
        End If
            
    Loop
    
    rs_tt_line.Close
    
    ArchExp.Close
    
Fin:
    'Cierro y libero todo
    If rs_field_value.State = adStateOpen Then rs_field_value.Close
    Set rs_field_value = Nothing
    If rs_topic.State = adStateOpen Then rs_topic.Close
    Set rs_topic = Nothing
    If rs_topic_field.State = adStateOpen Then rs_topic_field.Close
    Set rs_topic_field = Nothing
    If rs_field_value2.State = adStateOpen Then rs_field_value2.Close
    Set rs_field_value2 = Nothing
    If rs_intconfgen.State = adStateOpen Then rs_intconfgen.Close
    Set rs_intconfgen = Nothing
    If rs_interfaz.State = adStateOpen Then rs_interfaz.Close
    Set rs_interfaz = Nothing
    If rs_tt_topicline.State = adStateOpen Then rs_tt_topicline.Close
    Set rs_tt_topicline = Nothing
    If rs_tt_line.State = adStateOpen Then rs_tt_line.Close
    Set rs_tt_line = Nothing

Exit Sub

ME_Local:
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 1) & "---------------------------------------------"
    Flog.writeline Espacios(Tabulador * 1) & "Error: " & Err.Description
    Flog.writeline Espacios(Tabulador * 1) & "Ultimo SQL: " & StrSql
    Flog.writeline Espacios(Tabulador * 1) & "---------------------------------------------"
    Flog.writeline
End Sub

Function NewLine(ByVal linea As String, ByRef nro As Integer)
    StrSql = "INSERT INTO tt_line (line_number, line_text) VALUES (" & nro & ",'" & linea & "')"
    objConn.Execute StrSql, , adExecuteNoRecords
    
    nro = nro + 1
End Function

Sub imprimirTexto(Texto, archivo, Longitud, derecha)
'Rutina genérica para imprimir un TEXTO, de una LONGITUD determinada.
'Los sobrantes se rellenan con CARACTER

Dim cadena
Dim txt
Dim u
Dim longTexto
    
    If IsNull(Texto) Then
        longTexto = 0
        cadena = " "
    Else
        longTexto = Len(Texto)
        cadena = CStr(Texto)
    End If
    
    u = CInt(Longitud) - longTexto
    If u < 0 Then
        archivo.Write cadena
    Else
        If derecha Then
            archivo.Write cadena & String(u, " ")
        Else
            archivo.Write String(u, " ") & cadena
        End If
    End If

End Sub
