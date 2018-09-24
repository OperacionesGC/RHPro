Attribute VB_Name = "repResEmbargos"
Option Explicit

'Global Const Version = "1.01"
'Global Const FechaModificacion = "20/04/2007"
'Global Const UltimaModificacion = " " ' Martin Ferraro - Version Inicial
                                      
'Global Const Version = "1.02"
'Global Const FechaModificacion = "25/09/2007"
'Global Const UltimaModificacion = " " ' Martin Ferraro - DESDE HASTA NO OBLIG
                              
Global Const Version = "1.03"
Global Const FechaModificacion = "31/07/2009"
Global Const UltimaModificacion = " " ' Martin Ferraro - Encriptacion de string connection

Dim fs, f
'Global Flog

Dim NroProceso As Long

Global Path As String
Global HuboErrores As Boolean
Global EmpErrores As Boolean

Global titulofiltro As String
Global filtro As String
Global fecestr As String
Global tenro1  As Long
Global estrnro1  As Long
Global tenro2  As Long
Global estrnro2 As Long
Global tenro3 As Long
Global estrnro3 As Long
Global Orden As String
Global fec_desde As String
Global fec_hasta As String
Global terape As String
Global terape2 As String
Global ternom As String
Global ternom2 As String
Global embnro As Long
Global tpenro As Long
Global tpedesc As String
Global embest As String
Global mesini As Long
Global anioini As Long
Global mesfin As Long
Global aniofin As Long
Global Monto As Double
Global desc As Double
Global empleg As Long
Global Anio_desde As String
Global Anio_hasta As String
Global Mes_desde As String
Global Mes_hasta As String
Global estruc1 As Long
Global estruc2 As Long
Global estruc3 As Long


Private Sub Main()

Dim strCmdLine As String
Dim Nombre_Arch As String

Dim StrSql As String
Dim StrSql2 As String
Dim objRs As New ADODB.Recordset
Dim objRs2 As New ADODB.Recordset
Dim rsCuotas As New ADODB.Recordset
Dim I
Dim totalAcum
Dim cantRegistros
Dim PID As String
Dim parametros As String
Dim ArrParametros
Dim fecAuxHasta
Dim fecAuxDesde
Dim auxDesc As Double
Dim auxEmbimp As Double

'    strCmdLine = Command()
'    ArrParametros = Split(strCmdLine, " ", -1)
'    If UBound(ArrParametros) > 0 Then
'        If IsNumeric(ArrParametros(0)) Then
'            NroProceso = ArrParametros(0)
'            Etiqueta = ArrParametros(1)
'        Else
'            Exit Sub
'        End If
'    Else
'        If IsNumeric(strCmdLine) Then
'            NroProceso = strCmdLine
'        Else
'            Exit Sub
'        End If
'    End If

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
    
    On Error GoTo CE
    
    Nombre_Arch = PathFLog & "ReporteResumenEmbargos" & "-" & NroProceso & ".log"
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Nombre_Arch, True)
    
    Flog.writeline "_________________________________________________________________"
    Flog.writeline "Version = " & Version
    Flog.writeline "Modificacion = " & UltimaModificacion
    Flog.writeline "Fecha = " & FechaModificacion
    Flog.writeline "_________________________________________________________________"
    Flog.writeline

    TiempoInicialProceso = GetTickCount
    
    On Error Resume Next
    OpenConnection strconexion, objConn
    If Err.Number <> 0 Or Error_Encrypt Then
        Flog.writeline Espacios(Tabulador * 0) & "Problemas en la conexion"
        Exit Sub
    End If
    
    On Error GoTo CE
    HuboErrores = False
    
   
    Flog.writeline "Inicio Proceso de Reporte Resumen de Embargos : " & Now
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
       
       Flog.writeline Espacios(Tabulador * 0) & "Recuperando Parametros."
       'Obtengo los parametros del proceso
       parametros = objRs2!bprcparam
              
       If Not IsNull(parametros) Then
           ArrParametros = Split(parametros, "@")
           If UBound(ArrParametros) = 13 Then
           
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
                Orden = ArrParametros(9)
                ' Año desde del embargo
                Anio_desde = IIf(EsNulo(ArrParametros(10)), "", ArrParametros(10))
                ' Mes desde del embargo
                Mes_desde = IIf(EsNulo(ArrParametros(11)), "", ArrParametros(11))
                ' Año hasta del embargo
                Anio_hasta = IIf(EsNulo(ArrParametros(12)), "", ArrParametros(12))
                ' Mes hasta del embargo
                Mes_hasta = IIf(EsNulo(ArrParametros(13)), "", ArrParametros(13))
               
                
                '______________________________________________________
                Flog.writeline
                Flog.writeline Espacios(Tabulador * 1) & " Datos del Filtro: "
                Flog.writeline Espacios(Tabulador * 1) & "    Filtro: " & ArrParametros(1)
                Flog.writeline Espacios(Tabulador * 1) & "    Fecha Estr: " & ArrParametros(2)
                Flog.writeline Espacios(Tabulador * 1) & "    Estructuras: " & ArrParametros(3) & " - " & ArrParametros(4) & " - " & ArrParametros(5) & " - " & ArrParametros(6) & " - " & ArrParametros(7) & " - " & ArrParametros(8)
                Flog.writeline Espacios(Tabulador * 1) & "    Orden: " & ArrParametros(9)
                Flog.writeline Espacios(Tabulador * 1) & "    Fechas del Emb.: " & ArrParametros(11) & " - " & ArrParametros(10) & " al  " & ArrParametros(13) & " - " & ArrParametros(12)
                Flog.writeline
             
           Else
                Flog.writeline Espacios(Tabulador * 1) & "ERROR. La cantidad de parametros no es la esperada."
                Exit Sub
           End If
           
       Else
            Flog.writeline Espacios(Tabulador * 1) & "ERROR. No se encuentran los paramentros."
            Exit Sub
       End If
        
        'EMPIEZA EL PROCESO
        '------------------------------------------------------------------------------------------------------------------------
        'Control que la fecha inicio del embargo se encuentre dentro del rango de mes y anio del filtro
        '------------------------------------------------------------------------------------------------------------------------
        'Control desde
        StrSql2 = ""
        If Not EsNulo(Anio_desde) Then
            StrSql2 = " ( ( " & CLng(Anio_desde) & " < embargo.embanioini ) "
            StrSql2 = StrSql2 & " OR "
            StrSql2 = StrSql2 & " ( ( " & CLng(Anio_desde) & " = embargo.embanioini ) AND ( " & CLng(Mes_desde) & " <= embargo.embmesini ) ) )"
        End If
        'Control Hasta
        If Not EsNulo(Anio_hasta) Then
            If Len(StrSql2) <> 0 Then
                StrSql2 = StrSql2 & " AND "
            End If
            StrSql2 = StrSql2 & " ( ( embargo.embanioini < " & CLng(Anio_hasta) & " ) "
            StrSql2 = StrSql2 & " OR "
            StrSql2 = StrSql2 & " ( ( " & CLng(Anio_hasta) & " = embargo.embanioini ) AND ( embargo.embmesini <= " & CLng(Mes_hasta) & " ) ) )"
        End If
        '------------------------------------------------------------------------------------------------------------------------
        'GENERO LA StrSql QUE BUSCA LOS EMPLEADOS
        '------------------------------------------------------------------------------------------------------------------------

        If tenro3 <> 0 Then ' esto ocurre solo cuando se seleccionan los tres niveles
                    StrSql = " SELECT DISTINCT tipoemb.*, embargo.*,empleado.ternro,empleg, terape, terape2, ternom, ternom2 "
                    StrSql = StrSql & ", estact1.tenro tenro1, estact1.estrnro estrnro1 "
                    StrSql = StrSql & ", estact2.tenro tenro2, estact2.estrnro estrnro2 "
                    StrSql = StrSql & ", estact3.tenro tenro3, estact3.estrnro estrnro3 "
                    StrSql = StrSql & " FROM empleado  "
                    StrSql = StrSql & " INNER JOIN embargo ON empleado.ternro = embargo.ternro "
                    'StrSql = StrSql & " INNER JOIN embcuota ON embargo.embnro = embcuota.embnro "
                    StrSql = StrSql & " INNER JOIN tipoemb ON tipoemb.tpenro = embargo.tpenro "
                    StrSql = StrSql & " INNER JOIN his_estructura estact1 ON empleado.ternro = estact1.ternro  AND estact1.tenro  = " & tenro1
                    StrSql = StrSql & " AND (estact1.htetdesde<=" & ConvFecha(fecestr) & " AND (estact1.htethasta is null or estact1.htethasta>=" & ConvFecha(fecestr) & "))"
                    If estrnro1 <> 0 Then 'cuando se le asigna un valor al nivel 1
                            StrSql = StrSql & " AND estact1.estrnro =" & estrnro1
                    End If
                    StrSql = StrSql & " INNER JOIN his_estructura estact2 ON empleado.ternro = estact2.ternro  AND estact2.tenro  = " & tenro2
                    StrSql = StrSql & " AND (estact2.htetdesde<=" & ConvFecha(fecestr) & " AND (estact2.htethasta is null or estact2.htethasta>=" & ConvFecha(fecestr) & "))"
                    If estrnro2 <> 0 Then 'cuando se le asigna un valor al nivel 2
                        StrSql = StrSql & " AND estact2.estrnro =" & estrnro2
                    End If
                    StrSql = StrSql & " INNER JOIN his_estructura estact3 ON empleado.ternro = estact3.ternro  AND estact3.tenro  = " & tenro3
                    StrSql = StrSql & " AND (estact3.htetdesde<=" & ConvFecha(fecestr) & " AND (estact3.htethasta is null or estact3.htethasta>=" & ConvFecha(fecestr) & "))"
                    If estrnro3 <> 0 Then ' cuando se le asigna un valor al nivel 3
                            StrSql = StrSql & " AND estact3.estrnro =" & estrnro3
                    End If
                    StrSql = StrSql & " WHERE " & filtro
                    If Len(StrSql2) <> 0 Then
                        StrSql = StrSql & " AND (" & StrSql2 & ") "
                    End If
                    StrSql = StrSql & " ORDER BY tenro1,estrnro1,tenro2,estrnro2,tenro3,estrnro3," & Orden

        ElseIf tenro2 <> 0 Then  'ocurre cuando se selecciono hasta el segundo nivel
                    StrSql = "SELECT DISTINCT tipoemb.*, embargo.*,empleado.ternro,empleg, terape, terape2, ternom, ternom2"
                    StrSql = StrSql & ", estact1.tenro tenro1, estact1.estrnro estrnro1 "
                    StrSql = StrSql & ", estact2.tenro tenro2, estact2.estrnro estrnro2 "
                    StrSql = StrSql & " FROM empleado  "
                    StrSql = StrSql & " INNER JOIN embargo ON empleado.ternro = embargo.ternro "
                    'StrSql = StrSql & " INNER JOIN embcuota ON embargo.embnro = embcuota.embnro "
                    StrSql = StrSql & " INNER JOIN tipoemb ON tipoemb.tpenro = embargo.tpenro "
                    StrSql = StrSql & " INNER JOIN his_estructura estact1 ON empleado.ternro = estact1.ternro  AND estact1.tenro  = " & tenro1
                    StrSql = StrSql & " AND (estact1.htetdesde<=" & ConvFecha(fecestr) & " AND (estact1.htethasta is null or estact1.htethasta>=" & ConvFecha(fecestr) & "))"
                    If estrnro1 <> 0 Then
                        StrSql = StrSql & " AND estact1.estrnro =" & estrnro1
                    End If
                    StrSql = StrSql & " INNER JOIN his_estructura estact2 ON empleado.ternro = estact2.ternro  AND estact2.tenro  = " & tenro2
                    StrSql = StrSql & " AND (estact2.htetdesde<=" & ConvFecha(fecestr) & " AND (estact2.htethasta is null or estact2.htethasta>=" & ConvFecha(fecestr) & "))"
                    If estrnro2 <> 0 Then
                        StrSql = StrSql & " AND estact2.estrnro =" & estrnro2
                    End If
                    StrSql = StrSql & " WHERE " & filtro
                    If Len(StrSql2) <> 0 Then
                        StrSql = StrSql & " AND (" & StrSql2 & ") "
                    End If
                    StrSql = StrSql & " ORDER BY tenro1,estrnro1,tenro2,estrnro2," & Orden
           
        ElseIf tenro1 <> 0 Then  ' Cuando solo selecionamos el primer nivel
                    StrSql = "SELECT DISTINCT tipoemb.*, embargo.*,empleado.ternro,empleg, terape, terape2, ternom, ternom2 "
                    StrSql = StrSql & ", estact1.tenro tenro1, estact1.estrnro estrnro1 "
                    StrSql = StrSql & " FROM empleado  "
                    StrSql = StrSql & " INNER JOIN embargo ON empleado.ternro = embargo.ternro "
                    'StrSql = StrSql & " INNER JOIN embcuota ON embargo.embnro = embcuota.embnro "
                    StrSql = StrSql & " INNER JOIN tipoemb ON tipoemb.tpenro = embargo.tpenro "
                    StrSql = StrSql & " INNER JOIN his_estructura estact1 ON empleado.ternro = estact1.ternro  AND estact1.tenro  = " & tenro1
                    StrSql = StrSql & " AND (estact1.htetdesde<=" & ConvFecha(fecestr) & " AND (estact1.htethasta is null or estact1.htethasta>=" & ConvFecha(fecestr) & "))"
                    If estrnro1 <> 0 Then
                        StrSql = StrSql & " AND estact1.estrnro =" & estrnro1
                    End If
                    StrSql = StrSql & " WHERE " & filtro
                    If Len(StrSql2) <> 0 Then
                        StrSql = StrSql & " AND (" & StrSql2 & ") "
                    End If
                    StrSql = StrSql & " ORDER BY tenro1,estrnro1," & Orden
        
        Else  ' cuando no hay nivel de estructura seleccionado
                    StrSql = " SELECT DISTINCT tipoemb.*, embargo.*,empleado.ternro,empleg, terape, terape2, ternom, ternom2 "
                    StrSql = StrSql & " FROM empleado  "
                    StrSql = StrSql & " INNER JOIN embargo ON empleado.ternro = embargo.ternro "
                    'StrSql = StrSql & " INNER JOIN embcuota ON embargo.embnro = embcuota.embnro "
                    StrSql = StrSql & " INNER JOIN tipoemb ON tipoemb.tpenro = embargo.tpenro "
                    StrSql = StrSql & " WHERE  " & filtro
                    If Len(StrSql2) <> 0 Then
                        StrSql = StrSql & " AND (" & StrSql2 & ") "
                    End If
                    StrSql = StrSql & " ORDER BY " & Orden
        End If
                      
       'Busco el periodo desde
       OpenRecordset StrSql, objRs
        
        ' _________________________________________________________________________
        Flog.writeline "  SQL para control de los empleados periodo de las cuotas. "
        Flog.writeline "    " & StrSql
        Flog.writeline " "
'
       If objRs.EOF Then
          Flog.writeline "No se encontraron embargos para el Reporte."
          StrSql = "UPDATE batch_proceso SET  bprcprogreso =100, bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado' WHERE bpronro = " & NroProceso
          objConn.Execute StrSql, , adExecuteNoRecords
          Exit Sub

       Else
                
               'Obtengo los acumuladores sobre los que deberia generar la comparación
               
                cantRegistros = CLng(objRs.RecordCount)
                I = 1
               
               ' Genero los datos
               Do Until objRs.EOF
        
                    EmpErrores = False
                    embnro = objRs!embnro
                    
                    ' Genero los datos del embargo
                    Flog.writeline "Generando datos del embargo " & embnro
                                        
                    StrSql = " SELECT embccancela, embcimp, embcimpreal FROM embcuota "
                    StrSql = StrSql & " WHERE embnro = " & embnro
                    OpenRecordset StrSql, rsCuotas
                
                    auxDesc = 0
                    Do While Not rsCuotas.EOF
                       If rsCuotas!embccancela < 0 Then
                           auxDesc = auxDesc + CDbl(rsCuotas!embcimpreal)
                       End If
                
                        rsCuotas.MoveNext
                    Loop
                    rsCuotas.Close
                    
                    empleg = CLng(objRs!empleg)
                    
                    terape = CStr(objRs!terape)
                    terape2 = IIf(EsNulo(objRs!terape2), "", objRs!terape2)
                    ternom = CStr(objRs!ternom)
                    ternom2 = IIf(EsNulo(objRs!ternom2), "", objRs!ternom2)
                    tpenro = CLng(objRs!tpenro)
                    If tpenro = 0 Then
                        tpedesc = "Todos"
                    Else
                        tpedesc = CStr(objRs!tpedesabr)
                    End If
                    embest = CStr(objRs!embest)
                    mesini = IIf(EsNulo(objRs!embmesini), 0, objRs!embmesini)
                    anioini = IIf(EsNulo(objRs!embanioini), 0, objRs!embanioini)
                    mesfin = IIf(EsNulo(objRs!embmesfin), 0, objRs!embmesfin)
                    aniofin = IIf(EsNulo(objRs!embaniofin), 0, objRs!embaniofin)
                    desc = FormatNumber(CDbl(auxDesc), 2)
                    auxEmbimp = IIf(EsNulo(objRs!embimp), 0, objRs!embimp)
                    
                    If (CDbl(auxEmbimp) = 0) And (CDbl(auxDesc) > 0) Then
                        Monto = desc
                    Else
                        Monto = FormatNumber(CDbl(auxEmbimp), 2)
                    End If
                    
                    estruc1 = 0
                    estruc2 = 0
                    estruc3 = 0
                    If tenro1 <> 0 Then
                        estruc1 = IIf(EsNulo(objRs!estrnro1), 0, objRs!estrnro1)
                    End If
                    If tenro2 <> 0 Then
                        estruc2 = IIf(EsNulo(objRs!estrnro2), 0, objRs!estrnro2)
                    End If
                    If tenro3 <> 0 Then
                        estruc3 = IIf(EsNulo(objRs!estrnro3), 0, objRs!estrnro3)
                    End If
                                        
                    Flog.writeline " Insertando datos en la tabla "
                    ' Inserto los datos del detalle en la tabla
                    Call InsertarDatosDet
               
                    TiempoAcumulado = GetTickCount
                      
                    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Fix((I / cantRegistros) * 100) & _
                             ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'" & _
                             " WHERE bpronro = " & NroProceso
                    objConn.Execute StrSql, , adExecuteNoRecords
                     
                    I = I + 1
                                        
                    objRs.MoveNext
               Loop
           
       End If
    
    Else

       Exit Sub

    End If
    
    ' Insertar Datos Comunes de los embargos
    Call InsertarDatos(cantRegistros)

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
    Flog.writeline " Error: " & Err.Description
    Flog.writeline " SQL Ejecutado: " & StrSql


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
    StrSql = "INSERT INTO rep_res_emb (bpronro,aniodesde,aniohasta,mesdesde,meshasta,fechorarep,cant,titrep)"
    StrSql = StrSql & " VALUES ( "
    StrSql = StrSql & NroProceso & ","
    StrSql = StrSql & IIf(EsNulo(Anio_desde), 0, Anio_desde) & ","
    StrSql = StrSql & IIf(EsNulo(Anio_hasta), 0, Anio_hasta) & ","
    StrSql = StrSql & IIf(EsNulo(Mes_desde), 0, Mes_desde) & ","
    StrSql = StrSql & IIf(EsNulo(Mes_hasta), 0, Mes_hasta) & ","
    StrSql = StrSql & ConvFecha(Date) & "," & Cantidad & ",'" & titulofiltro & "')"
    objConn.Execute StrSql, , adExecuteNoRecords

    Exit Sub
                
MError:
    Flog.writeline "    Error: " & Err.Description
    Flog.writeline "    SQL Ejecutado: " & StrSql
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
    StrSql = "INSERT INTO rep_res_emb_det (bpronro,embnro,tpedesabr,embest,embmesini,embanioini," & _
             "embmesfin,embaniofin,empleg,terape,ternom2,terape2,ternom,embimp,embimpdesc," & _
             "tenro1,estrnro1,tenro2,estrnro2,tenro3,estrnro3,fecestr) "
    StrSql = StrSql & "VALUES (" & _
             NroProceso & "," & embnro & ",'" & tpedesc & "','" & embest & "'," & mesini & _
             "," & anioini & "," & mesfin & "," & aniofin & "," & empleg & ",'" & _
             terape & "','" & ternom2 & "','" & terape2 & "','" & ternom & "'," & _
             Monto & "," & desc & "," & tenro1 & "," & estruc1 & "," & tenro2 & "," & estruc2 & "," & _
             tenro3 & "," & estruc3 & "," & ConvFecha(fecestr) & ")"
    objConn.Execute StrSql, , adExecuteNoRecords
        
    Exit Sub
                
MError2:
    Flog.writeline "    Error: " & Err.Description
    Flog.writeline "    SQL Ejecutado: " & StrSql
    HuboErrores = True
    EmpErrores = True
    Exit Sub

End Sub


