Attribute VB_Name = "repCuotasEmbargosF"
Option Explicit

Dim fs, f
Global Flog

Dim NroProceso As Long

Global Path As String
Global HuboErrores As Boolean
Global EmpErrores As Boolean

Global pronro2 As String
Global prodesc2 As String

Global titulofiltro As String
Global titulo_rep_hist As String
Global filtro As String
Global fecestr As String
Global tenro1  As Long
Global estrnro1  As Long
Global tenro2  As Long
Global estrnro2 As Long
Global tenro3 As Long
Global estrnro3 As Long
Global orden As String
Global procesos As String
Global terape As String
Global terape2 As String
Global ternom As String
Global ternom2 As String
Global embnro As Long
Global tpenro As Long
Global tpedesc As String
Global embest As String
Global embcnro As Long
Global embcimp As Double
Global desc As Double
Global empleg As Long



Private Sub Main()

Dim strCmdLine As String
Dim Nombre_Arch As String

Dim StrSql As String
Dim StrSql2 As String
Dim objRs As New ADODB.Recordset
Dim objRs2 As New ADODB.Recordset
Dim rsProcesos As New ADODB.Recordset
Dim i As Integer
Dim totalAcum
Dim cantRegistros
Dim PID As String
Dim parametros As String
Dim ArrParametros
Dim procesos_sep
Dim long_tit As Integer
Dim HayDatos As Boolean


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

    On Error GoTo CE
    
    TiempoInicialProceso = GetTickCount
    OpenConnection strconexion, objConn
    
    HuboErrores = False
    HayDatos = True
    
    Nombre_Arch = PathFLog & "ReporteCuotaEmbargos" & "-" & NroProceso & ".log"
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Nombre_Arch, True)
    
    Flog.writeline "Inicio Proceso de Reporte de Cuotas de Embargos : " & Now
    Flog.writeline "Cambio el estado del proceso a Procesando"
    
    ' Obtengo el Process ID
    PID = GetCurrentProcessId
    Flog.writeline "PID = " & PID
    
    'Cambio el estado del proceso a Procesando
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcestado = 'Procesando', bprcpid = " & PID & ", bprcprogreso = 0 WHERE bpronro = " & NroProceso
    objConn.Execute StrSql, , adExecuteNoRecords
    
    Flog.writeline "Obtengo los datos del proceso"
    
    TiempoAcumulado = GetTickCount
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE bpronro = " & NroProceso
    OpenRecordset StrSql, objRs2
    
    If Not objRs2.EOF Then
       
       'Obtengo los parametros del proceso
       parametros = objRs2!bprcparam
       ArrParametros = Split(parametros, "@")
              
       'Obtengo el Titulo
        titulofiltro = ArrParametros(0)
        
       'Obtengo el filtro utilizado como restricciones a la busqueda de embargos
        filtro = ArrParametros(1)
        If InStr(filtro, "embargo.embest") > 0 Then
            filtro = Replace(filtro, "embargo.embest = A", "embargo.embest = 'A'")
            filtro = Replace(filtro, "embargo.embest = E", "embargo.embest = 'E'")
            filtro = Replace(filtro, "embargo.embest = F", "embargo.embest = 'F'")
            filtro = Replace(filtro, "embargo.embest = I", "embargo.embest = 'I'")
        End If
        
        ' Fecha a considerar las estructuras
        fecestr = CDate(ArrParametros(2))
        
        ' Nro tercero de la primer estructura
        tenro1 = CLng(ArrParametros(3))
        
        'Codigo de la primer estructura
        estrnro1 = CLng(ArrParametros(4))
        
        ' Nro tercero de la segunda estructura
        tenro2 = CLng(ArrParametros(5))
        
        ' Codigo de la segunda estructura
        estrnro2 = CLng(ArrParametros(6))
        
        ' Nro de tercero de la tercer estructura
        tenro3 = CLng(ArrParametros(7))
        
        ' Codigo de la tercer estructura
        estrnro3 = CLng(ArrParametros(8))

        ' String conteniendo el orden en el cual se debe realizar la busqueda de embargos
        orden = ArrParametros(9)

        ' String conteniendo los procesos de los cuales se listaran las cuotas
        procesos = ArrParametros(10)
              

        'EMPIEZA EL PROCESO

        '------------------------------------------------------------------------------------------------------------------------
        'GENERO LA SQL QUE BUSCA EL PERIODOS
        '------------------------------------------------------------------------------------------------------------------------

        procesos_sep = Split(procesos, ",", -1, 1)

        i = 0
        StrSql2 = " (embcuota.pronro IN (" & procesos & "))"
        StrSql2 = StrSql2 & " AND (embcuota.embccancela < 0)"
        
        ' armo el titulo del reporte historico
        titulo_rep_hist = CStr(Date) & " - Procesos: "
        
        Do While i <= UBound(procesos_sep)
            
            ' Armo el titulo historico
            StrSql = " SELECT prodesc FROM proceso WHERE pronro = " & CLng(procesos_sep(i))
            OpenRecordset StrSql, rsProcesos
            
            If Not rsProcesos.EOF Then
                titulo_rep_hist = titulo_rep_hist & CStr(rsProcesos!prodesc) & " - "
            End If
            
            rsProcesos.Close
            
            i = i + 1
            
        Loop
        
        ' Termino de armar el titulo historico
        long_tit = Len(titulo_rep_hist) - 3
        titulo_rep_hist = Mid(titulo_rep_hist, 1, long_tit)
        If long_tit > 200 Then
            titulo_rep_hist = Mid(titulo_rep_hist, 1, 196) & "..."
        End If
        
        '------------------------------------------------------------------------------------------------------------------------
        'GENERO LA StrSql QUE BUSCA LOS EMPLEADOS
        '------------------------------------------------------------------------------------------------------------------------

        If tenro3 <> 0 Then ' esto ocurre solo cuando se seleccionan los tres niveles
                    StrSql = " SELECT DISTINCT proceso.prodesc, embcuota.*, tipoemb.*, embargo.*,v_empleado.ternro,empleg, terape, terape2, ternom, ternom2 "
                    StrSql = StrSql & ", estact1.tenro AS tenro1, estact1.estrnro AS estrnro1 "
                    StrSql = StrSql & ", estact2.tenro AS tenro2, estact2.estrnro AS estrnro2 "
                    StrSql = StrSql & ", estact3.tenro AS tenro3, estact3.estrnro AS estrnro3 "
                    StrSql = StrSql & " FROM v_empleado  "
                    StrSql = StrSql & " INNER JOIN embargo ON v_empleado.ternro = embargo.ternro "
                    StrSql = StrSql & " INNER JOIN embcuota ON embargo.embnro = embcuota.embnro "
                    StrSql = StrSql & " INNER JOIN tipoemb ON tipoemb.tpenro = embargo.tpenro "
                    StrSql = StrSql & " INNER JOIN proceso ON embcuota.pronro = proceso.pronro "
                    StrSql = StrSql & " INNER JOIN his_estructura estact1 ON v_empleado.ternro = estact1.ternro  AND estact1.tenro  = " & tenro1
                    StrSql = StrSql & " AND (estact1.htetdesde<=" & ConvFecha(fecestr) & " AND (estact1.htethasta is null or estact1.htethasta>=" & ConvFecha(fecestr) & "))"
                    If estrnro1 <> 0 Then 'cuando se le asigna un valor al nivel 1
                            StrSql = StrSql & " AND estact1.estrnro =" & estrnro1
                    End If
                    StrSql = StrSql & " INNER JOIN his_estructura estact2 ON v_empleado.ternro = estact2.ternro  AND estact2.tenro  = " & tenro2
                    StrSql = StrSql & " AND (estact2.htetdesde<=" & ConvFecha(fecestr) & " AND (estact2.htethasta is null or estact2.htethasta>=" & ConvFecha(fecestr) & "))"
                    If estrnro2 <> 0 Then 'cuando se le asigna un valor al nivel 2
                        StrSql = StrSql & " AND estact2.estrnro =" & estrnro2
                    End If
                    StrSql = StrSql & " INNER JOIN his_estructura estact3 ON v_empleado.ternro = estact3.ternro  AND estact3.tenro  = " & tenro3
                    StrSql = StrSql & " AND (estact3.htetdesde<=" & ConvFecha(fecestr) & " AND (estact3.htethasta is null or estact3.htethasta>=" & ConvFecha(fecestr) & "))"
                    If estrnro3 <> 0 Then ' cuando se le asigna un valor al nivel 3
                            StrSql = StrSql & " AND estact3.estrnro =" & estrnro3
                    End If
                    StrSql = StrSql & " WHERE " & filtro
                    StrSql = StrSql & " AND (" & StrSql2 & ") "
                    StrSql = StrSql & " ORDER BY tenro1,estrnro1,tenro2,estrnro2,tenro3,estrnro3," & orden

        ElseIf tenro2 <> 0 Then  'ocurre cuando se selecciono hasta el segundo nivel
                    StrSql = "SELECT DISTINCT proceso.prodesc, embcuota.*,tipoemb.*, embargo.*,v_empleado.ternro,empleg, terape, terape2, ternom, ternom2"
                    StrSql = StrSql & ", estact1.tenro AS tenro1, estact1.estrnro AS estrnro1 "
                    StrSql = StrSql & ", estact2.tenro AS tenro2, estact2.estrnro AS estrnro2 "
                    StrSql = StrSql & " FROM v_empleado  "
                    StrSql = StrSql & " INNER JOIN embargo ON v_empleado.ternro = embargo.ternro "
                    StrSql = StrSql & " INNER JOIN embcuota ON embargo.embnro = embcuota.embnro "
                    StrSql = StrSql & " INNER JOIN tipoemb ON tipoemb.tpenro = embargo.tpenro "
                    StrSql = StrSql & " INNER JOIN proceso ON embcuota.pronro = proceso.pronro "
                    StrSql = StrSql & " INNER JOIN his_estructura estact1 ON v_empleado.ternro = estact1.ternro  AND estact1.tenro  = " & tenro1
                    StrSql = StrSql & " AND (estact1.htetdesde<=" & ConvFecha(fecestr) & " AND (estact1.htethasta is null or estact1.htethasta>=" & ConvFecha(fecestr) & "))"
                    If estrnro1 <> 0 Then
                        StrSql = StrSql & " AND estact1.estrnro =" & estrnro1
                    End If
                    StrSql = StrSql & " INNER JOIN his_estructura estact2 ON v_empleado.ternro = estact2.ternro  AND estact2.tenro  = " & tenro2
                    StrSql = StrSql & " AND (estact2.htetdesde<=" & ConvFecha(fecestr) & " AND (estact2.htethasta is null or estact2.htethasta>=" & ConvFecha(fecestr) & "))"
                    If estrnro2 <> 0 Then
                        StrSql = StrSql & " AND estact2.estrnro =" & estrnro2
                    End If
                    StrSql = StrSql & " WHERE " & filtro
                    StrSql = StrSql & " AND (" & StrSql2 & ") "
                    StrSql = StrSql & " ORDER BY tenro1,estrnro1,tenro2,estrnro2," & orden
           
        ElseIf tenro1 <> 0 Then  ' Cuando solo selecionamos el primer nivel
                    StrSql = "SELECT DISTINCT proceso.prodesc, embcuota.*,tipoemb.*, embargo.*,v_empleado.ternro,empleg, terape, terape2, ternom, ternom2 "
                    StrSql = StrSql & ", estact1.tenro AS tenro1, estact1.estrnro AS estrnro1 "
                    StrSql = StrSql & " FROM v_empleado  "
                    StrSql = StrSql & " INNER JOIN embargo ON v_empleado.ternro = embargo.ternro "
                    StrSql = StrSql & " INNER JOIN embcuota ON embargo.embnro = embcuota.embnro "
                    StrSql = StrSql & " INNER JOIN tipoemb ON tipoemb.tpenro = embargo.tpenro "
                    StrSql = StrSql & " INNER JOIN proceso ON embcuota.pronro = proceso.pronro "
                    StrSql = StrSql & " INNER JOIN his_estructura estact1 ON v_empleado.ternro = estact1.ternro  AND estact1.tenro  = " & tenro1
                    StrSql = StrSql & " AND (estact1.htetdesde<=" & ConvFecha(fecestr) & " AND (estact1.htethasta is null or estact1.htethasta>=" & ConvFecha(fecestr) & "))"
                    If estrnro1 <> 0 Then
                        StrSql = StrSql & " AND estact1.estrnro =" & estrnro1
                    End If
                    StrSql = StrSql & " WHERE " & filtro
                    StrSql = StrSql & " AND (" & StrSql2 & ") "
                    StrSql = StrSql & " ORDER BY tenro1,estrnro1," & orden
        
        Else  ' cuando no hay nivel de estructura seleccionado
                    StrSql = " SELECT DISTINCT proceso.prodesc, embcuota.*,tipoemb.*, embargo.*,v_empleado.ternro,empleg, terape, terape2, ternom, ternom2 "
                    StrSql = StrSql & " FROM v_empleado  "
                    StrSql = StrSql & " INNER JOIN embargo ON v_empleado.ternro = embargo.ternro "
                    StrSql = StrSql & " INNER JOIN embcuota ON embargo.embnro = embcuota.embnro "
                    StrSql = StrSql & " INNER JOIN tipoemb ON tipoemb.tpenro = embargo.tpenro "
                    StrSql = StrSql & " INNER JOIN proceso ON embcuota.pronro = proceso.pronro "
                    StrSql = StrSql & " WHERE " & filtro
                    StrSql = StrSql & " AND (" & StrSql2 & ") "
                    StrSql = StrSql & " ORDER BY " & orden
        End If
                      
       'Busco el periodo desde
       OpenRecordset StrSql, objRs
        
       If objRs.EOF Then
          Flog.writeline "No se encontraron cuotas canceladas de embargos para el Reporte."
          HayDatos = False
               
       Else
                              
                cantRegistros = CLng(objRs.RecordCount)
                i = 1
               
               ' Genero los datos
               Do Until objRs.EOF
        
                    EmpErrores = False
                    embnro = objRs!embnro
                    
                    ' Genero los datos del embargo
                    Flog.writeline "Generando datos del embargo " & embnro
                                        
                    empleg = CLng(objRs!empleg)
                    
                    terape = CStr(objRs!terape)
                    terape2 = IIf(objRs!terape2 <> Null, objRs!terape2, "")
                    ternom = CStr(objRs!ternom)
                    ternom2 = IIf(objRs!ternom2 <> Null, objRs!terape2, "")
                    tpenro = CLng(objRs!tpenro)
                    If tpenro = 0 Then
                            tpedesc = "Todos"
                          Else
                            tpedesc = CStr(objRs!tpedesabr)
                    End If
                    embest = CStr(objRs!embest)
                    embcnro = objRs!embcnro
                    embcimp = FormatNumber(CDbl(objRs!embcimp), 2)
                    desc = FormatNumber(CDbl(objRs!embcimpreal), 2)
                    If tenro1 <> 0 Then
                        estrnro1 = objRs!estrnro1
                    End If
                    If tenro2 <> 0 Then
                        estrnro2 = objRs!estrnro2
                    End If
                    If tenro3 <> 0 Then
                        estrnro3 = objRs!estrnro3
                    End If
                    ' Modif. pedidas por Javier - GdeCos - 30/05/2005
                    ' Nro. Proceso y Descripcion
                    pronro2 = objRs!pronro
                    prodesc2 = objRs!prodesc
                    
                    Flog.writeline "Insertando datos en la tabla "
                    
                    ' Inserto los datos del detalle en la tabla
                    Call InsertarDatosDet
               
                    TiempoAcumulado = GetTickCount
                      
                    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Fix((i / cantRegistros) * 100) & _
                             ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'" & _
                             " WHERE bpronro = " & NroProceso
                    objConn.Execute StrSql, , adExecuteNoRecords
                     
                    i = i + 1
                                        
                    objRs.MoveNext
               Loop
               
       End If
    
    Else

       Exit Sub

    End If
    
    ' Insertar Datos Comunes de los embargos
    If HayDatos Then
        Call InsertarDatos(cantRegistros)
    End If

    'Actualizo el estado del proceso
    If Not HuboErrores Then
       StrSql = "UPDATE batch_proceso SET  bprcprogreso =100, bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado' WHERE bpronro = " & NroProceso
       Flog.writeline "Proceso Finalizado Correctamente"
    Else
       StrSql = "UPDATE batch_proceso SET  bprcprogreso =100, bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Incompleto' WHERE bpronro = " & NroProceso
       Flog.writeline "Proceso Incompleto"
    End If
    
    objConn.Execute StrSql, , adExecuteNoRecords

    Flog.writeline "Fin :" & Now
    Flog.Close
    If objRs.State = adStateOpen Then objRs.Close

    Exit Sub
    
CE:
    HuboErrores = True
    Flog.writeline " Error: " & Err.Description & Now

End Sub

'------------------------------------------------------------------------------------
' Se encarga de Insertar los datos comunes de la consulta en la tabla de Resultados
'------------------------------------------------------------------------------------
Sub InsertarDatos(ByVal Cantidad As Integer)

    Dim StrSql As String
    
    On Error GoTo MError
    
    '-------------------------------------------------------------------------------
    'Inserto los datos en la BD
    '-------------------------------------------------------------------------------
    StrSql = "INSERT INTO rep_cuota_emb (bpronro,fechorarep,cant,titrep,rep_hist_tit) "
    StrSql = StrSql & "VALUES (" & _
             NroProceso & "," & ConvFecha(Date) & "," & Cantidad & ",'" & titulofiltro & _
             "','" & titulo_rep_hist & "')"
    objConn.Execute StrSql, , adExecuteNoRecords
        
    Exit Sub
                
MError:
    Flog.writeline " Error: " & Err.Description
    HuboErrores = True
    EmpErrores = True
    Exit Sub

End Sub

'------------------------------------------------------------------------------------
' Se encarga de Insertar los datos de la consulta en la tabla de Resultados
'------------------------------------------------------------------------------------
Sub InsertarDatosDet()

    Dim StrSql As String
    
    On Error GoTo MError2
    
    '-------------------------------------------------------------------------------
    'Inserto los datos en la BD
    '-------------------------------------------------------------------------------
    StrSql = "INSERT INTO rep_cuota_emb_det (bpronro,embnro,tpedesabr,embest,embcnro,embcimp," & _
             "embcimpdesc,empleg,terape,ternom2,terape2,ternom," & _
             "tenro1,estrnro1,tenro2,estrnro2,tenro3,estrnro3,fecestr,pronro,prodesc) "
    StrSql = StrSql & "VALUES (" & _
             NroProceso & "," & embnro & ",'" & tpedesc & "','" & embest & "'," & embcnro & _
             "," & embcimp & "," & desc & "," & empleg & ",'" & _
             terape & "','" & ternom2 & "','" & terape2 & "','" & ternom & "'," & _
             tenro1 & "," & estrnro1 & "," & tenro2 & "," & estrnro2 & "," & _
             tenro3 & "," & estrnro3 & "," & ConvFecha(fecestr) & "," & pronro2 & ",'" & prodesc2 & "')"
    objConn.Execute StrSql, , adExecuteNoRecords
        
    Exit Sub
                
MError2:
    Flog.writeline " Error: " & Err.Description
    HuboErrores = True
    EmpErrores = True
    Exit Sub

End Sub


