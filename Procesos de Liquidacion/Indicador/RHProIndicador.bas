Attribute VB_Name = "RHProIndicador"
Option Explicit

Dim HuboErrores As Boolean
Dim objRs2 As New ADODB.Recordset
Dim l_sql As String
Dim l_schednro As Integer
Dim l_dia As String
Dim l_Hora As String
Dim TituloAlerta As String
Dim dirsalidas As String
Dim usuario As String
Global Contador As Integer
Dim l_indsqlcorrida As String
Global sqlEnOracle As Boolean
Global Progreso As Double
Global IncPorc As Double
Global Cantidad_de_indicadores As Long



Function ReplaceFields(ByVal query As String, ByVal aleconsnro As String, ByVal aleNro As String) As String
    Dim sql As String
    Dim Campo As String
    
    Do Until InStr(1, query, "@@") = 0
        Campo = Mid(query, InStr(1, query, "@@") + 2, Len(query))
        Campo = Mid(Campo, 1, InStr(1, Campo, "@@") - 1)
        sql = "select alepa_valor, alepa_tipo from ale_param where alenro = " & aleNro
        sql = sql & " and aleconsnro = " & aleconsnro
        sql = sql & " and upper(alepa_nombre) = '" & UCase(Campo) & "'"
        OpenRecordset sql, objRs2
        If Not objRs2.EOF Then
            If UCase(objRs2!alepa_tipo) = "D" Or UCase(objRs2!alepa_tipo) = "S" Then
              query = Replace(query, "@@" & Campo & "@@", "'" & objRs2!alepa_valor & "'")
            Else
              query = Replace(query, "@@" & Campo & "@@", objRs2!alepa_valor)
            End If
        Else
            Flog.writeline "Error en campo parametro '" & Campo & "'"
            Exit Function
        End If
        If objRs2.State = adStateOpen Then objRs2.Close
        
 
    Loop
    ReplaceFields = query
End Function

Sub Main()

Dim objRs As New ADODB.Recordset
Dim objRs2 As New ADODB.Recordset
Dim objRsInd As New ADODB.Recordset
Dim objconn2 As New ADODB.Connection
Dim rs_batch_proceso As New ADODB.Recordset

Dim strCmdLine As String
Dim Nombre_Arch As String
Dim StrSql As String
Dim PID As String
Dim I As Integer
Dim arr
Dim aleNro As String
Dim AlertaAccion As Integer
Dim sqlale As String
Dim DispararAlerta As Boolean
Dim NroProceso As String
Dim colcount As Integer
Dim Mensaje As String
Dim auxi As String
Dim tipoMails As Integer
Dim ArrParametros
Dim l_indnro As Integer
Dim l_Resu As Single
Dim l_Resu_Total As Single
Dim Ind_Detallado
Dim l_Ter
Dim l_Fecha As String
Dim l_Hora As String
'Dim Manual As String
Dim Manual As Boolean

'FGZ - 01/11/2013 ----------------
Dim rs_his As New ADODB.Recordset
Dim QuedaronParametros As Boolean
'FGZ - 01/11/2013 ----------------


' JPB - 06-11-2013 ---
Dim l_indhisdesabr As String
Dim incaux As Double
' JPB - 06-11-2013 --

'strCmdLine = Command()
'ArrParametros = Split(strCmdLine, " ", -1)
'If UBound(ArrParametros) > 0 Then
'    If IsNumeric(ArrParametros(0)) Then
'        NroProcesoBatch = ArrParametros(0)
'        Etiqueta = ArrParametros(1)
'    Else
'        Exit Sub
'    End If
'Else
'    If IsNumeric(strCmdLine) Then
'        NroProcesoBatch = strCmdLine
'    Else
'        Exit Sub
'    End If
'End If

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





On Error GoTo ME_Main
    
    HuboErrores = False
    
    
    ' carga las configuraciones basicas, formato de fecha, string de conexion,
    ' tipo de BD y ubicacion del archivo de log
    Call CargarConfiguracionesBasicas
    
    
    'Nombre_Arch = PathFLog & "Indicador" & Day(Date) & "-" & Int(Month(Date)) & "-" & Year(Date) & "_" & Left(Time, 2) & Mid(Time, 4, 2) & ".log"
    Nombre_Arch = PathFLog & "Indicador" & "-" & NroProcesoBatch & ".log"
    
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
    Flog.writeline
    Flog.writeline "Inicio Indicador : " & Now

    On Error Resume Next
    'Abro la conexion
    OpenConnection strconexion, objConn
    If Err.Number <> 0 Or Error_Encrypt Then
        Flog.writeline "Problemas en la conexion"
        'Exit Sub
        GoTo CierroLog
    End If
    OpenConnection strconexion, objconnProgreso
    If Err.Number <> 0 Or Error_Encrypt Then
        Flog.writeline "Problemas en la conexion"
        'Exit Sub
        GoTo CierroLog
    End If
    
    On Error GoTo CE
    
    On Error GoTo 0 ' desactivo el manejador de errores que exista al momento
    
    'FGZ - 06/08/2012 --------- Control de versiones ------
    Version_Valida = ValidarV(Version, 192, TipoBD)
    If Not Version_Valida Then
        'Actualizo el progreso
        MyBeginTrans
            StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcprogreso = 0, bprcestado = 'Error de Version', bprcpid = " & PID & " WHERE bpronro = " & NroProcesoBatch
            objconnProgreso.Execute StrSql, , adExecuteNoRecords
        MyCommitTrans
        GoTo CierroLog
    End If
    'FGZ - 06/08/2012 --------- Control de versiones ------
    
    On Error GoTo ErrorSQL
    
    'Cambio el estado del proceso a Procesando y el PID
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcprogreso = 0, bprcestado = 'Procesando', bprcpid = " & PID & " WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE btprcnro = 192 AND bpronro =" & NroProcesoBatch
    OpenRecordset StrSql, rs_batch_proceso
    If rs_batch_proceso.EOF Then
        Flog.writeline "No existe el proceso " & NroProcesoBatch
        'FGZ 06/08/2012 --------------------
        'Exit Sub
        GoTo Fin
        'FGZ 06/08/2012 --------------------
    Else
'        arr = Split(rs_batch_proceso!bprcparam, " ")
'        Manual = False
'        If (UBound(arr) - LBound(arr) + 1) > 1 Then
'            If Not EsNulo(arr(0)) Then
'                Manual = CBool(arr(0))
'            End If
'        End If
        Flog.writeline "Parametros del Indicador: " & rs_batch_proceso!bprcparam
        If Not EsNulo(rs_batch_proceso!bprcparam) Then
            If rs_batch_proceso!bprcparam <> True And rs_batch_proceso!bprcparam <> False Then
                Flog.writeline "    El parametro no es valido (true o false)"
                'FGZ 06/08/2012 --------------------
                'Exit Sub
                GoTo Fin
                'FGZ 06/08/2012 --------------------
            Else
                Manual = CBool(rs_batch_proceso!bprcparam)
                Flog.writeline "    Ejecucion Manual"
            End If
            
        End If
    End If
    
    'Verifica si indicadores esta activo y si debe correr las sql de Oracle o no
    sqlEnOracle = False
    StrSql = "SELECT *"
    StrSql = StrSql & " FROM confper"
    StrSql = StrSql & " WHERE confnro = 2 AND confactivo = -1"
    OpenRecordset StrSql, objRs
    If objRs.EOF Then
        Flog.writeline "Indicadores no se encuentra Activo"
        GoTo Fin
    Else
        If Not EsNulo(objRs!confint) Then
            sqlEnOracle = (CLng(objRs!confint) = -1)
            Flog.writeline "SQL en ORACLE = " & sqlEnOracle
        End If
    End If
    Flog.writeline
    
    ' Para Cada Indicador Activo
     
       StrSql = "SELECT count(indnro) cantidad FROM indicador WHERE indactivo = -1"
       OpenRecordset StrSql, objRs
       
        TiempoAcumulado = GetTickCount
        Cantidad_de_indicadores = objRs!Cantidad
         IncPorc = 100 / Cantidad_de_indicadores
        
    
    StrSql = "SELECT * FROM indicador WHERE indactivo = -1"
    OpenRecordset StrSql, objRs
    If objRs.EOF Then
        Flog.writeline "No hay ningun Indicador Activo."
        'FGZ 06/08/2012 --------------------
        'Exit Sub
        GoTo Fin
        'FGZ 06/08/2012 --------------------
    End If
    l_Fecha = Day(Date) & "/" & Int(Month(Date)) & "/" & Year(Date)
    l_Fecha = ConvFecha(l_Fecha)
    'l_Hora = Left(Time, 2) & Mid(Time, 4, 2)
    l_Hora = Format_StrNro(Hour(Time), 2, True, 0) & Format_StrNro(Minute(Time), 2, True, 0)
    
        
    Flog.writeline
    Do Until objRs.EOF
'        Progreso = Progreso + IncPorc

        Flog.writeline
        Flog.writeline "Indicador " & objRs!indnro & " - " & objRs!indnom & " scheduler " & objRs!schednro
        l_schednro = objRs!schednro
        l_indnro = objRs!indnro
        
        l_indsqlcorrida = ""
        If Not IsNull(objRs!inddetalle) Then
             Ind_Detallado = CBool(objRs!inddetalle)
        Else
            Ind_Detallado = False
        End If
        'Obtengo la sql dependiendo si es oracle o no
        If sqlEnOracle Then
            If EsNulo(objRs!indsqlcorridaora) Then l_indsqlcorrida = "" Else l_indsqlcorrida = objRs!indsqlcorridaora
        Else
            If EsNulo(objRs!indsqlcorrida) Then l_indsqlcorrida = "" Else l_indsqlcorrida = objRs!indsqlcorrida
        End If


        If Not EsNulo(l_indsqlcorrida) Then
            'Calcular Proxima Ejecucion
            'Si ahora ent
            'If EjecutarAhora(l_indnro, l_schednro) Then
            If EjecutarAhora(l_indnro, l_schednro) Or Manual Then
                'Obtener SQL
                ReemplazarParametros
                
                'FGZ - 01/11/2013 ----------------------------------
                QuedaronParametros = InStr(1, l_indsqlcorrida, "[") <> 0
                If Not QuedaronParametros Then
                
                    ReemplazarFunciones (l_indnro)
                    'Ejecutar SQL
                    Flog.writeline Espacios(Tabulador * 1) & "Se ejecuta el siguiente SQL: " & l_indsqlcorrida
                    OpenRecordset l_indsqlcorrida, objRs2
                    Flog.writeline "l_indsqlcorrida:" & l_indsqlcorrida
                     Flog.writeline "cantidad de resultados:" & objRs2.RecordCount
                     If objRs2.RecordCount > 0 Then
                        incaux = IncPorc / objRs2.RecordCount
                     Else
                        incaux = IncPorc
                    End If
                     
                    'Insertar Resultado
                    If objRs2.EOF Then
                        Flog.writeline "No hay resultado."
                    Else
                        If Ind_Detallado Then
                        Flog.writeline "Hay resultado."
                        'JPB - En el caso que el indicador tenga Detalle por empleado se debe cargar el resultado del query en la tabla ind_historia_det
                        'Para que dicha query funcione deberá recuperar dos campos:
                        'El pimero el valor del tercero asociado, y el segundo debe traer un valor numérico
                          
                          'JPB - 06-11-201-  Recupero la descripcion del indicador ----------
                            l_indhisdesabr = ""
                            
                            StrSql = "select inddesext from indicador  "
                            StrSql = StrSql & " WHERE indnro = " & l_indnro
                            OpenRecordset StrSql, objRsInd

                            If Not objRsInd.EOF Then
                                 If Not EsNulo(objRsInd!inddesext) Then
                                    l_indhisdesabr = objRsInd!inddesext
                                 End If
                            End If
                          ' -----------------------------------------------------------------
                            
                            
                            'Incorporo el detalle de la query para cada tercero, en la tabla ind_historia_det
                            Do While Not objRs2.EOF
                              
                                If objRs2.Fields.Count > 1 Then
                                    l_Ter = objRs2(0)
                                    l_Resu = objRs2(1)
                                 
                                                        
                                'FGZ - 01/11/2013 --------------------------------------------------------------------------------
                                'Controlo por el indice unico que el registro no exista
                                StrSql = "SELECT * FROM ind_historia_det "
                                StrSql = StrSql & " WHERE indnro = " & l_indnro
                                StrSql = StrSql & " AND indhisfec = " & l_Fecha
                                StrSql = StrSql & " AND indhishora = '" & l_Hora & "'"
                                StrSql = StrSql & " AND ternro = " & l_Ter
                               
                                OpenRecordset StrSql, rs_his
                                   
                                If rs_his.EOF Then
                                
                                    'JPB - 06-11-2013 - Agrego como descripcion la del indicador
                                    'StrSql = "INSERT INTO ind_historia_det (indnro, indhisfec, indhishora,ternro,indhisvalor) VALUES ("
                                    StrSql = "INSERT INTO ind_historia_det (indnro, indhisfec, indhishora,ternro,indhisvalor,indhisdesabr) VALUES ("
                                    StrSql = StrSql & l_indnro & ","
                                    StrSql = StrSql & l_Fecha & ",'"
                                    StrSql = StrSql & l_Hora & "',"
                                    StrSql = StrSql & l_Ter & ","
                                    'StrSql = StrSql & l_Resu & ")"
                                    StrSql = StrSql & l_Resu & ","
                                    StrSql = StrSql & "'" & Left(l_indhisdesabr, 50) & "')"
                                Else
                                    StrSql = "UPDATE ind_historia_det SET indhisvalor = " & l_Resu
                                    StrSql = StrSql & " WHERE indnro = " & l_indnro
                                    StrSql = StrSql & " AND indhisfec = " & l_Fecha
                                    StrSql = StrSql & " AND indhishora = '" & l_Hora & "'"
                                    StrSql = StrSql & " AND ternro = " & l_Ter
                                  
                                End If
                                   
                                objConn.Execute StrSql, , adExecuteNoRecords
                                   
                                'FGZ - 01/11/2013 --------------------------------------------------------------------------------
                                l_Resu_Total = l_Resu_Total + l_Resu
                                
    '                            StrSql = "INSERT INTO ind_historia_det (indnro, indhisfec, indhishora,ternro,indhisvalor) VALUES ("
    '                            StrSql = StrSql & l_indnro & ","
    '                            StrSql = StrSql & l_Fecha & ",'"
    '                            StrSql = StrSql & l_Hora & "',"
    '                            StrSql = StrSql & l_Ter & ","
    '                            StrSql = StrSql & l_Resu & ")"
    '                            objConn.Execute StrSql, , adExecuteNoRecords
    '
    '                            l_Resu_Total = l_Resu_Total + l_Resu
                             
                            Else
                                   Flog.writeline " Error. El indicador no es detallado"

                                 End If
                                objRs2.MoveNext
                                
                                    Progreso = Progreso + incaux
                           
                                   StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso
                                    StrSql = StrSql & ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso))
                                    StrSql = StrSql & "' WHERE bpronro = " & NroProcesoBatch
                                    objconnProgreso.Execute StrSql, , adExecuteNoRecords
                                   ' Flog.writeline "  actualizar progreso: " & StrSql
                           
                            Loop
                            'l_resu = objRs2(0)
                            
                            
                            'FGZ - 01/11/2013 --------------------------------------------------------------------------------
                            'Controlo por el indice unico que el registro no exista
                            StrSql = "SELECT * FROM ind_historia "
                            StrSql = StrSql & " WHERE indnro = " & l_indnro
                            StrSql = StrSql & " AND indhisfec = " & l_Fecha
                            StrSql = StrSql & " AND indhishora = '" & l_Hora & "'"
                            OpenRecordset StrSql, rs_his
                            If rs_his.EOF Then
                                StrSql = "INSERT INTO ind_historia (indnro, indhisfec, indhishora,indhisvalor) VALUES ("
                                StrSql = StrSql & l_indnro & ","
                                StrSql = StrSql & l_Fecha & ",'"
                                StrSql = StrSql & l_Hora & "',"
                                'FGZ - 15/01/2013 ----------------------
                                'StrSql = StrSql & l_Resu & ")"
                                StrSql = StrSql & l_Resu_Total & ")"
                                'FGZ - 15/01/2013 ----------------------
                            Else
                                StrSql = "UPDATE ind_historia SET indhisvalor = " & l_Resu_Total
                                StrSql = StrSql & " WHERE indnro = " & l_indnro
                                StrSql = StrSql & " AND indhisfec = " & l_Fecha
                                StrSql = StrSql & " AND indhishora = '" & l_Hora & "'"
                            End If
                            objConn.Execute StrSql, , adExecuteNoRecords
                            'FGZ - 01/11/2013 --------------------------------------------------------------------------------
                        Else
                            Flog.writeline "Hay Resultado. si re"
                            l_Resu = objRs2(0)
                            
                            
                            'FGZ - 01/11/2013 --------------------------------------------------------------------------------
                            'Controlo por el indice unico que el registro no exista
                            StrSql = "SELECT * FROM ind_historia "
                            StrSql = StrSql & " WHERE indnro = " & l_indnro
                            StrSql = StrSql & " AND indhisfec = " & l_Fecha
                            StrSql = StrSql & " AND indhishora = '" & l_Hora & "'"
                            OpenRecordset StrSql, rs_his
                            If rs_his.EOF Then
                                StrSql = "INSERT INTO ind_historia (indnro, indhisfec, indhishora,indhisvalor) VALUES ("
                                StrSql = StrSql & l_indnro & ","
                                StrSql = StrSql & l_Fecha & ",'"
                                StrSql = StrSql & l_Hora & "',"
                                StrSql = StrSql & l_Resu & ")"
                            Else
                                StrSql = "UPDATE ind_historia SET indhisvalor = " & l_Resu
                                StrSql = StrSql & " WHERE indnro = " & l_indnro
                                StrSql = StrSql & " AND indhisfec = " & l_Fecha
                                StrSql = StrSql & " AND indhishora = '" & l_Hora & "'"
                            End If
                            objConn.Execute StrSql, , adExecuteNoRecords
                            'FGZ - 01/11/2013 --------------------------------------------------------------------------------
                            
                        End If
                    End If
                    'Actualizar Ultima Ejecucion
                Else
                    'Quedaron parametros
                    Flog.writeline "    Quedaron parametros sin resolver. Revise la configuracion del indicador " & l_indnro
                
                End If
            Else
                Flog.writeline Espacios(Tabulador * 1) & "Por planificacion no corresponde ejecucion."
            End If
        Else
            Flog.writeline Espacios(Tabulador * 1) & "SQL NULA."
        End If
      
Terminar:
        If objRs2.State = adStateOpen Then objRs2.Close
        objRs.MoveNext
        
        'StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Error' WHERE bpronro = " & NroProcesoBatch
        
       ' objconnProgreso.Execute StrSql, , adExecuteNoRecords

        
    Loop

Fin:
    If Not HuboError Then
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado' ,bprcprogreso=100 WHERE bpronro = " & NroProcesoBatch
    Else
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Error' WHERE bpronro = " & NroProcesoBatch
    End If
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
CierroLog:
    Flog.writeline
    Flog.writeline "Fin procesamiento " & Now
    Flog.writeline "-----------------------------------------------------------------"
    Flog.Close
    
    If objconnProgreso.State = adStateOpen Then objconnProgreso.Close
    If objConn.State = adStateOpen Then objConn.Close
    If objRs.State = adStateOpen Then objRs.Close
    
    Exit Sub

CE:
    HuboErrores = True
    Flog.writeline " Error: " & Err.Description & Now
    Flog.writeline " Indicador Nro: " & l_indnro
    Flog.writeline " SQL Indicador: " & l_indsqlcorrida
    Flog.writeline " SQL Sistema: " & StrSql
    

    
ErrorSQL:
    HuboErrores = True
    Flog.writeline " Error: " & Err.Description & Now
    Flog.writeline " Indicador Nro: " & l_indnro
    Flog.writeline " SQL Indicador: " & l_indsqlcorrida
    Flog.writeline " SQL Sistema: " & StrSql
    
    GoTo Terminar
    
ME_Main:
    'Error general

End Sub


Public Function EjecutarAhora(p_indnro As Integer, p_schednro As Integer) As Boolean
    Dim l_UltFecha
    Dim l_UltHora
    Dim l_frectipnro
    Dim l_schedhora
    Dim l_alesch_frecrep
    Dim StrSql As String
    
    
    'Consulto la fecha y hora de la ultima ejecution del Indicador
    'StrSql = "SELECT TOP 1 indhisfec, indhishora FROM ind_historia WHERE indnro = " & p_indnro
    StrSql = "SELECT indhisfec, indhishora FROM ind_historia WHERE indnro = " & p_indnro
    StrSql = StrSql & " ORDER BY indhisfec DESC, indhishora DESC "
       
    If objRs2.State = adStateOpen Then objRs2.Close
    OpenRecordset StrSql, objRs2
    'Si nunca se ejecuto entonces Ejecutar Ahora!
    If objRs2.EOF Then
        EjecutarAhora = True
        Exit Function
    End If
    'Sino comparar la fecha de ultima ejecucion
    'contra la fecha actual y analizar si corresponde ejecutarlo ahora
    
    l_UltFecha = objRs2("indhisfec")
    l_UltHora = objRs2("indhishora")
    
    'Busco el schedule asociado al Indicador
    
    l_sql = "SELECT frectipnro, alesch_fecini, schedhora, alesch_frecrep, alesch_fecfin "
    l_sql = l_sql & "FROM ale_sched "
    l_sql = l_sql & "WHERE schednro = " & p_schednro
  
    If objRs2.State = adStateOpen Then objRs2.Close
    OpenRecordset l_sql, objRs2
    
    '07/08/2007 - Martin Ferraro - No preguntaban If objRs2.EOF y daba error si habian borrado el schedule
    If objRs2.EOF Then
        
        'El schedule no se encuentra
        EjecutarAhora = False
    
    Else
    
        l_frectipnro = objRs2!frectipnro
        l_schedhora = objRs2!schedhora
        l_alesch_frecrep = objRs2!alesch_frecrep
    
       ' Si se ejecuta Diariamiente
       If l_frectipnro = 1 Then
           'Se ejecuta si la fecha de ultima corrida es previa a hoy
           'y la hora actual es mayor a la programada
           'EjecutarAhora = DateValue(l_UltFecha) < DateValue(Date) And _
                   Int(Left(Time, 2) & Mid(Time, 4, 2)) + 1 > _
                   Int(Left(l_schedhora, 2) & Mid(l_schedhora, 4, 2))
           
           'Se chequea que la hora llegue con el separado ":" y se agregan 0 a la izquierda
           If InStr(l_schedhora, ":") > 0 Then
                EjecutarAhora = DateValue(l_UltFecha) < DateValue(Date) Or _
                        (DateValue(l_UltFecha) = DateValue(Date) And _
                        Int(Hour(Now) & Format_StrNro(Minute(Now), 2, True, 0)) + 1 > _
                        Int(Left(l_schedhora, 2) & Mid(l_schedhora, 4, 2)))
           Else
                EjecutarAhora = DateValue(l_UltFecha) < DateValue(Date) Or _
                        (DateValue(l_UltFecha) = DateValue(Date) And _
                        Int(Hour(Now) & Format_StrNro(Minute(Now), 2, True, 0)) + 1 > _
                        Int(Left(l_schedhora, 2) & Mid(l_schedhora, 3, 2)))
           End If
           Flog.writeline "         Ultima ejecucion: " & l_UltFecha & ", dia que se intento ejecutar: " & Date
           Flog.writeline "         Son las " & Int(Hour(Now) & Format_StrNro(Minute(Now), 2, True, 0)) + 1 & " el indicador esta planificado para las " & l_schedhora
           Exit Function
       End If
       ' Si se ejecuta Mensualmente
       If l_frectipnro = 3 Then
           'Se ejecuta si el mes de ultima corrida es previo a hoy,
           'el dia de hoy es mayor al dia programado
           'y la hora actual es mayor a la programada
           'EjecutarAhora = (Year(l_UltFecha) < Year(Date) Or Month(l_UltFecha) < Month(Date)) And _
                   Day(Date) >= Int(l_alesch_frecrep) And _
                   Int(Left(Time, 2) & Mid(Time, 4, 2)) + 1 > _
                   Int(Left(l_schedhora, 2) & Mid(l_schedhora, 4, 2))
                   
           'Se chequea que la hora llegue con el separado ":" y se agregan 0 a la izquierda
           If InStr(l_schedhora, ":") > 0 Then
                EjecutarAhora = (Year(l_UltFecha) < Year(Date) Or Month(l_UltFecha) < Month(Date)) Or _
                        ((Year(l_UltFecha) = Year(Date) Or Month(l_UltFecha) = Month(Date)) And _
                        Int(Hour(Now) & Format_StrNro(Minute(Now), 2, True, 0)) + 1 > _
                        Int(Left(l_schedhora, 2) & Mid(l_schedhora, 4, 2)))
           Else
                EjecutarAhora = (Year(l_UltFecha) < Year(Date) Or Month(l_UltFecha) < Month(Date)) Or _
                        ((Year(l_UltFecha) = Year(Date) Or Month(l_UltFecha) = Month(Date)) And _
                        Int(Hour(Now) & Format_StrNro(Minute(Now), 2, True, 0)) + 1 > _
                        Int(Left(l_schedhora, 2) & Mid(l_schedhora, 3, 2)))
           End If
           Flog.writeline "         Ultima ejecucion: " & l_UltFecha & ", dia que se intento ejecutar: " & Date
           Flog.writeline "         Son las " & Int(Hour(Now) & Format_StrNro(Minute(Now), 2, True, 0)) + 1 & " el indicador esta planificado para las " & l_schedhora

           Exit Function
       End If
       ' Si se ejecuta Semanalmente
       If l_frectipnro = 2 Then
           'Se ejecuta si fecha de ultima corrida es menor que hoy,
           'el dia de la semana actual es el dia de la semana programada
           'y la hora actual es mayor a la programada
           'EjecutarAhora = DateValue(l_UltFecha) < DateValue(Date) And _
                       Weekday(Date) = l_alesch_frecrep And _
                       Int(Left(Time, 2) & Mid(Time, 4, 2)) + 1 > _
                       Int(Left(l_schedhora, 2) & Mid(l_schedhora, 4, 2))
           
           'Se chequea que la hora llegue con el separado ":" y se agregan 0 a la izquierda
           If InStr(l_schedhora, ":") > 0 Then
                EjecutarAhora = DateValue(l_UltFecha) < DateValue(Date) Or _
                        (DateValue(l_UltFecha) = DateValue(Date) And _
                        Int(Hour(Now) & Format_StrNro(Minute(Now), 2, True, 0)) + 1 > _
                        Int(Left(l_schedhora, 2) & Mid(l_schedhora, 4, 2)))
           Else
                EjecutarAhora = DateValue(l_UltFecha) < DateValue(Date) Or _
                        (DateValue(l_UltFecha) = DateValue(Date) And _
                        Int(Hour(Now) & Format_StrNro(Minute(Now), 2, True, 0)) + 1 > _
                        Int(Left(l_schedhora, 2) & Mid(l_schedhora, 3, 2)))
           End If
           Flog.writeline "         Ultima ejecucion: " & l_UltFecha & ", dia que se intento ejecutar: " & Date
           Flog.writeline "         Son las " & Int(Hour(Now) & Format_StrNro(Minute(Now), 2, True, 0)) + 1 & " el indicador esta planificado para las " & l_schedhora

           Exit Function
       End If
    
       ' Si se ejecuta Temporalmente
       If l_frectipnro = 4 Then
           'Si la fecha actual es mayor que la de ultima corrida mas
           'los dias programados
           EjecutarAhora = DateValue(Now - l_alesch_frecrep) >= DateValue(l_UltFecha)
           Flog.writeline "         Ultima ejecucion: " & l_UltFecha & ", dia que se intento ejecutar con el corrimiento: " & DateValue(Now - l_alesch_frecrep)
           Exit Function
       End If
       
    End If
    If objRs2.State = adStateOpen Then objRs2.Close
End Function

Public Sub ReemplazarParametros()
    'Reemplazo cada codigo de variable por su nombre
    Dim l_sql As String
    Dim l_str As String
    
    l_sql = "SELECT * "
    l_sql = l_sql & " FROM indparametro "
    If objRs2.State = adStateOpen Then objRs2.Close
    OpenRecordset l_sql, objRs2
    Do Until objRs2.EOF
        l_str = "[" & objRs2("indparamnro") & "]"
        l_indsqlcorrida = Replace(l_indsqlcorrida, l_str, RTrim(objRs2("indparamvalor")))
        objRs2.MoveNext

    Loop
    
    objRs2.Close
End Sub

Public Sub ReemplazarFunciones(p_indnro As Integer)
    'Reemplazo cada codigo de variable por su nombre
    Dim l_sql As String
    Dim l_str As String
    
    l_str = "(0"
    
    l_sql = "SELECT estrnro "
    l_sql = l_sql & " FROM ind_limestr"
    l_sql = l_sql & " WHERE indnro = " & p_indnro
    If objRs2.State = adStateOpen Then objRs2.Close
    OpenRecordset l_sql, objRs2
    Do Until objRs2.EOF
        l_str = l_str & "," & objRs2("estrnro")
        objRs2.MoveNext
    Loop
    
    objRs2.Close

    l_str = l_str & ")"
    l_indsqlcorrida = Replace(l_indsqlcorrida, "fxEstructurasAsociadas()", l_str)


End Sub


Public Sub OpenRecordsetWithConn(strSQLQuery As String, ByRef objRs As ADODB.Recordset, ByRef Conn As ADODB.Connection, Optional lockType As LockTypeEnum = adLockReadOnly)
Dim pos1 As Integer
Dim pos2 As Integer
Dim aux As String

    'Abre un recordset con la consulta strSQLQuery
    If objRs.State <> adStateClosed Then
        If objRs.lockType <> adLockReadOnly Then objRs.UpdateBatch
        objRs.Close
    End If
    
    objRs.CacheSize = 500

    objRs.Open strSQLQuery, Conn, adOpenDynamic, lockType, adCmdText
    
End Sub


