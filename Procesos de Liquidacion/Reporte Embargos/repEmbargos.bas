Attribute VB_Name = "repEmbargos"
Option Explicit

'Version 1.00 al 30-05-2005

'Global Const Version = "1.01"
'Global Const FechaModificacion = "06/03/2007" ' Leticia Amadio
'Global Const UltimaModificacion = " " 'Sacar las vistas - agregar versión y comentarios
                                      
'Global Const Version = "1.02"
'Global Const FechaModificacion = "10/03/2007" ' Gustavo Ring
'Global Const UltimaModificacion = " " 'Se agregaron logs

'Global Const Version = "1.03" ' Cesar Stankunas
'Global Const FechaModificacion = "05/08/2009"
'Global Const UltimaModificacion = ""    'Encriptacion de string connection

Global Const Version = "1.04" ' Ruiz Miriam - CAS-27675 - VSO – Chacomer - Bug en reporte de embargo
Global Const FechaModificacion = "11/11/2014"
Global Const UltimaModificacion = ""    'Se hace un update a la tabla batch proceso en el caso que no haya embargos en el período seleccionado


' __________________________________________________________________________

Dim fs, f
'Global Flog

Dim NroProceso As Long

Global Path As String
Global HuboErrores As Boolean
Global EmpErrores As Boolean

Global pronro2 As String
Global listaacunro As String

Global titulofiltro As String
Global filtro As String
Global fecestr As String
Global tenro1  As Long
Global estrnro1  As Long
Global tenro2  As Long
Global estrnro2 As Long
Global tenro3 As Long
Global estrnro3 As Long
Global orden As String
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
Global mesini As Integer
Global anioini As Integer
Global mesfin As Integer
Global aniofin As Integer
Global Monto As Double
Global Desc As Double
Global empleg As Long



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

    Nombre_Arch = PathFLog & "ReporteEmbargos" & "-" & NroProceso & ".log"
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Nombre_Arch, True)
    
    Flog.writeline "_________________________________________________________________"
    Flog.writeline "Version = " & Version
    Flog.writeline "Modificacion = " & UltimaModificacion
    Flog.writeline "Fecha = " & FechaModificacion
    Flog.writeline "_________________________________________________________________"
    Flog.writeline
    
    ' Obtengo el Process ID
    PID = GetCurrentProcessId
    Flog.writeline "PID = " & PID
    
    TiempoInicialProceso = GetTickCount
    On Error Resume Next
    OpenConnection strconexion, objConn
    If Err.Number <> 0 Or Error_Encrypt Then
        Flog.writeline Espacios(Tabulador * 0) & "Problemas en la conexion"
        Exit Sub
    End If
    On Error GoTo CE
    HuboErrores = False
    
    Flog.writeline "Inicio Proceso de Reporte de Embargos : " & Now
    Flog.writeline "Cambio el estado del proceso a Procesando"
    
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

        ' Fecha inicial del periodo en el cual se deben buscar los embargos
        fec_desde = ArrParametros(10)
        
        ' Fecha final del periodo en el cual se deben buscar los embargos
        fec_hasta = ArrParametros(11)
       
        
        '______________________________________________________
        Flog.writeline " Datos del Filtro: "
        Flog.writeline "    Filtro: " & ArrParametros(1)
        Flog.writeline "    Fecha Estr: " & ArrParametros(2)
        Flog.writeline "    Estructuras: " & ArrParametros(3) & " - " & ArrParametros(4) & " - " & ArrParametros(5) & " - " & ArrParametros(6) & " - " & ArrParametros(7) & " - " & ArrParametros(8)
        Flog.writeline "    Orden: " & ArrParametros(9)
        Flog.writeline "    Período de Emb.: " & ArrParametros(10) & " - " & ArrParametros(11)
        Flog.writeline " "
         
        'EMPIEZA EL PROCESO

        '------------------------------------------------------------------------------------------------------------------------
        'GENERO LA SQL QUE BUSCA EL PERIODOS
        '------------------------------------------------------------------------------------------------------------------------

        fecAuxHasta = Split(fec_hasta, "/", -1, 1)
        fecAuxDesde = Split(fec_desde, "/", -1, 1)

        ' Controlo que el periodo tenga cuotas generadas en el rango dado
        StrSql2 = " ( ( (embcuota.embcanio > " & Int(fecAuxDesde(2)) & ") AND (embcuota.embcanio < " & Int(fecAuxHasta(2)) & ") ) "

        StrSql2 = StrSql2 & " OR "

        StrSql2 = StrSql2 & " ( (embcuota.embcanio = " & Int(fecAuxDesde(2)) & ") AND (embcuota.embcanio < " & Int(fecAuxHasta(2)) & ") AND (embcuota.embcmes >= " & Int(fecAuxDesde(1)) & ") ) "
        
        StrSql2 = StrSql2 & " OR "
        
        StrSql2 = StrSql2 & " ( (embcuota.embcanio > " & Int(fecAuxDesde(2)) & ") AND (embcuota.embcanio = " & Int(fecAuxHasta(2)) & ") AND (embcuota.embcmes <= " & Int(fecAuxHasta(1)) & ") ) "
        
        StrSql2 = StrSql2 & " OR "
        
        StrSql2 = StrSql2 & " ( (embcuota.embcanio = " & Int(fecAuxDesde(2)) & ") AND (embcuota.embcanio = " & Int(fecAuxHasta(2)) & ") AND (embcuota.embcmes >= " & Int(fecAuxDesde(1)) & ") AND (embcuota.embcmes <= " & Int(fecAuxHasta(1)) & ") ) )"
        
        ' ____________________________________________________________
        Flog.writeline "  SQL para control del periodo de las cuotas. "
        Flog.writeline "    " & StrSql2
        Flog.writeline " "
'        StrSql2 = " ( ( (embargo.embanioini > " & Int(fecAuxDesde(2)) & ") AND (embargo.embaniofin < " & Int(fecAuxHasta(2)) & ") ) "
'
'        StrSql2 = StrSql2 & " OR "
'
'        StrSql2 = StrSql2 & "( ( (" & Int(fecAuxDesde(2)) & " <= embargo.embanioini) AND (" & Int(fecAuxHasta(2)) & "> embargo.embaniofin ) )"
'        StrSql2 = StrSql2 & "   AND (" & Int(fecAuxDesde(1)) & " <= embargo.embmesini ) )"
'
'        StrSql2 = StrSql2 & " OR "
'
'        StrSql2 = StrSql2 & "( ( (" & Int(fecAuxDesde(2)) & " < embargo.embanioini) AND (" & Int(fecAuxHasta(2)) & ">= embargo.embaniofin ) )"
'        StrSql2 = StrSql2 & "   AND (" & Int(fecAuxHasta(1)) & " >= embargo.embmesfin ) )"
'
'        StrSql2 = StrSql2 & " OR "
'
'        StrSql2 = StrSql2 & "( (" & Int(fecAuxDesde(2)) & " = embargo.embanioini) AND (" & Int(fecAuxHasta(2)) & " = embargo.embaniofin) "
'        StrSql2 = StrSql2 & "   AND (embargo.embmesfin  <= " & Int(fecAuxHasta(1)) & ") AND (embargo.embmesini >= " & Int(fecAuxDesde(1)) & ") ) )"

        '------------------------------------------------------------------------------------------------------------------------
        'GENERO LA StrSql QUE BUSCA LOS EMPLEADOS
        '------------------------------------------------------------------------------------------------------------------------

        If tenro3 <> 0 Then ' esto ocurre solo cuando se seleccionan los tres niveles
                    StrSql = " SELECT DISTINCT tipoemb.*, embargo.*,empleado.ternro,empleg, terape, terape2, ternom, ternom2 "
                    StrSql = StrSql & ", estact1.tenro AS tenro1, estact1.estrnro AS estrnro1 "
                    StrSql = StrSql & ", estact2.tenro AS tenro2, estact2.estrnro AS estrnro2 "
                    StrSql = StrSql & ", estact3.tenro AS tenro3, estact3.estrnro AS estrnro3 "
                    StrSql = StrSql & " FROM empleado  "
                    StrSql = StrSql & " INNER JOIN embargo ON empleado.ternro = embargo.ternro "
                    StrSql = StrSql & " INNER JOIN embcuota ON embargo.embnro = embcuota.embnro "
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
                    StrSql = StrSql & " AND (" & StrSql2 & ") "
                    StrSql = StrSql & " ORDER BY tenro1,estrnro1,tenro2,estrnro2,tenro3,estrnro3," & orden

        ElseIf tenro2 <> 0 Then  'ocurre cuando se selecciono hasta el segundo nivel
                    StrSql = "SELECT DISTINCT tipoemb.*, embargo.*,empleado.ternro,empleg, terape, terape2, ternom, ternom2"
                    StrSql = StrSql & ", estact1.tenro AS tenro1, estact1.estrnro AS estrnro1 "
                    StrSql = StrSql & ", estact2.tenro AS tenro2, estact2.estrnro AS estrnro2 "
                    StrSql = StrSql & " FROM empleado  "
                    StrSql = StrSql & " INNER JOIN embargo ON empleado.ternro = embargo.ternro "
                    StrSql = StrSql & " INNER JOIN embcuota ON embargo.embnro = embcuota.embnro "
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
                    StrSql = StrSql & " AND (" & StrSql2 & ") "
                    StrSql = StrSql & " ORDER BY tenro1,estrnro1,tenro2,estrnro2," & orden
           
        ElseIf tenro1 <> 0 Then  ' Cuando solo selecionamos el primer nivel
                    StrSql = "SELECT DISTINCT tipoemb.*, embargo.*,empleado.ternro,empleg, terape, terape2, ternom, ternom2 "
                    StrSql = StrSql & ", estact1.tenro AS tenro1, estact1.estrnro AS estrnro1 "
                    StrSql = StrSql & " FROM empleado  "
                    StrSql = StrSql & " INNER JOIN embargo ON empleado.ternro = embargo.ternro "
                    StrSql = StrSql & " INNER JOIN embcuota ON embargo.embnro = embcuota.embnro "
                    StrSql = StrSql & " INNER JOIN tipoemb ON tipoemb.tpenro = embargo.tpenro "
                    StrSql = StrSql & " INNER JOIN his_estructura estact1 ON empleado.ternro = estact1.ternro  AND estact1.tenro  = " & tenro1
                    StrSql = StrSql & " AND (estact1.htetdesde<=" & ConvFecha(fecestr) & " AND (estact1.htethasta is null or estact1.htethasta>=" & ConvFecha(fecestr) & "))"
                    If estrnro1 <> 0 Then
                        StrSql = StrSql & " AND estact1.estrnro =" & estrnro1
                    End If
                    StrSql = StrSql & " WHERE " & filtro
                    StrSql = StrSql & " AND (" & StrSql2 & ") "
                    StrSql = StrSql & " ORDER BY tenro1,estrnro1," & orden
        
        Else  ' cuando no hay nivel de estructura seleccionado
                    StrSql = " SELECT DISTINCT tipoemb.*, embargo.*,empleado.ternro,empleg, terape, terape2, ternom, ternom2 "
                    StrSql = StrSql & " FROM empleado  "
                    StrSql = StrSql & " INNER JOIN embargo ON empleado.ternro = embargo.ternro "
                    StrSql = StrSql & " INNER JOIN embcuota ON embargo.embnro = embcuota.embnro "
                    StrSql = StrSql & " INNER JOIN tipoemb ON tipoemb.tpenro = embargo.tpenro "
                    StrSql = StrSql & " WHERE  " & filtro
                    StrSql = StrSql & " AND (" & StrSql2 & ") "
                    StrSql = StrSql & " ORDER BY " & orden
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
             StrSql = "UPDATE batch_proceso SET  bprcprogreso =100, bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Incompleto' WHERE bpronro = " & NroProceso
            Flog.writeline "Proceso Incompleto"
            objConn.Execute StrSql, , adExecuteNoRecords
            Flog.writeline "Fin :" & Now
             Flog.Close
          Exit Sub

       Else
                
               'Obtengo los acumuladores sobre los que deberia generar la comparación
'              Call CargarAcum(pronro1, pronro2, rsAcum)
               
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
                    mesini = CInt(objRs!embmesini)
                    anioini = CInt(objRs!embanioini)
                    mesfin = CInt(objRs!embmesfin)
                    aniofin = CInt(objRs!embaniofin)
                    Desc = FormatNumber(CDbl(auxDesc), 2)
                    If (CDbl(objRs!embimp) = 0) And (CDbl(auxDesc) > 0) Then
                        Monto = Desc
                    Else
                        Monto = FormatNumber(CDbl(objRs!embimp), 2)
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
    StrSql = "INSERT INTO rep_embargos (bpronro,fecdesde,fechasta,fechorarep,cant,titrep) "
    StrSql = StrSql & "VALUES (" & _
             NroProceso & "," & ConvFecha(fec_desde) & "," & ConvFecha(fec_hasta) & "," & ConvFecha(Date) & "," & Cantidad & _
             ",'" & titulofiltro & "')"
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
    StrSql = "INSERT INTO rep_embargos_det (bpronro,embnro,tpedesabr,embest,embmesini,embanioini," & _
             "embmesfin,embaniofin,empleg,terape,ternom2,terape2,ternom,embimp,embimpdesc," & _
             "tenro1,estrnro1,tenro2,estrnro2,tenro3,estrnro3,fecestr) "
    StrSql = StrSql & "VALUES (" & _
             NroProceso & "," & embnro & ",'" & tpedesc & "','" & embest & "'," & mesini & _
             "," & anioini & "," & mesfin & "," & aniofin & "," & empleg & ",'" & _
             terape & "','" & ternom2 & "','" & terape2 & "','" & ternom & "'," & _
             Monto & "," & Desc & "," & tenro1 & "," & estrnro1 & "," & tenro2 & "," & estrnro2 & "," & _
             tenro3 & "," & estrnro3 & "," & ConvFecha(fecestr) & ")"
    objConn.Execute StrSql, , adExecuteNoRecords
        
    Exit Sub
                
MError2:
    Flog.writeline "    Error: " & Err.Description
    Flog.writeline "    SQL Ejecutado: " & StrSql
    HuboErrores = True
    EmpErrores = True
    Exit Sub

End Sub


