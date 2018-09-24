Attribute VB_Name = "mdlRepAcumParcial"
Option Explicit

'---------------------------------------------------------------------------------------------------------
Const Version = "1.00"
Const FechaVersion = "28/10/2013"
'Sebastian Stremel - CAS-20908 - Sykes - Exportaciones GTI -> NOV. LIQ
'---------------------------------------------------------------------------------------------------------

Global HuboError As Boolean



Public Sub Main()
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento inicial del Generador de Reporte SIJP.
' Autor      : FGZ
' Fecha      : 20/01/2003
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim objconnMain As New ADODB.Connection
Dim strCmdLine
Dim Nombre_Arch As String
Dim rs_batch_proceso As New ADODB.Recordset
'Dim objConnProgreso As New ADODB.Connection
Dim PID As String
Dim bprcparam As String
Dim ArrParametros
Dim NroProcesoBatch
Dim TiempoInicialProceso
Dim TiempoFinalProceso


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
    
    ' carga las configuraciones basicas, formato de fecha, string de conexion,
    ' tipo de BD y ubicacion del archivo de log
    Call CargarConfiguracionesBasicas

    On Error GoTo ME_Main
    Nombre_Arch = PathFLog & "Generacion_Reporte_AcumuladoParcial" & "-" & NroProcesoBatch & ".log"
    
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
    'Abro la conexion
    OpenConnection strconexion, objConn
    If Err.Number <> 0 Then
        Flog.writeline "Problemas en la conexion"
        Exit Sub
    End If
    OpenConnection strconexion, objConnProgreso

    If Err.Number <> 0 Then
        Flog.writeline "Problemas en la conexion"
        Exit Sub
    End If
    On Error GoTo 0
    
    On Error GoTo ME_Main
    
    
    'Cambio el estado del proceso a Procesando
    'StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcprogreso = 0 ,bprcestado = 'Procesando', bprcpid = " & PID & " WHERE bpronro = " & NroProcesoBatch
    'objConnProgreso.Execute StrSql, , adExecuteNoRecords
    
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcprogreso = 0 ,bprcestado = 'Procesando', bprcpid = " & PID & " WHERE bpronro = " & NroProcesoBatch
    objConnProgreso.Execute StrSql, , adExecuteNoRecords
    
    Flog.writeline "Pone el estado en procesando"
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE btprcnro = 408 AND bpronro =" & NroProcesoBatch
    OpenRecordset StrSql, rs_batch_proceso
    
    TiempoInicialProceso = GetTickCount
    
    If Not rs_batch_proceso.EOF Then
        bprcparam = rs_batch_proceso!bprcparam
        rs_batch_proceso.Close
        Set rs_batch_proceso = Nothing
        Call RepAcumParcial(NroProcesoBatch, bprcparam)
    Else
        Flog.writeline "no encontró el proceso"
    End If
    
    TiempoFinalProceso = GetTickCount
    Flog.writeline "Tiempo del proceso (milisegundos): " & (TiempoFinalProceso - TiempoInicialProceso)
    
    If Not HuboError Then
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado' WHERE bpronro = " & NroProcesoBatch
    Else
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Error' WHERE bpronro = " & NroProcesoBatch
    End If
    objConnProgreso.Execute StrSql, , adExecuteNoRecords
    
Fin:
    Flog.Close
    'If objConn.State = adStateOpen Then objConn.Close
    If objConnProgreso.State = adStateOpen Then objConnProgreso.Close
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
        objConnProgreso.Execute StrSql, , adExecuteNoRecords
    MyCommitTrans
End Sub

Public Sub RepAcumParcial(ByVal NroProceso, ByVal parametros)

Dim StrSql As String
Dim sqlaux2 As String
Dim rs2 As New ADODB.Recordset
Dim rs3 As New ADODB.Recordset
Dim rs4 As New ADODB.Recordset
Dim StrSqlAux As String
Dim listaParametros
Dim filtro As String
Dim listaHoras As String
Dim tenro1 As String
Dim tenro2 As String
Dim tenro3 As String
Dim estrnro1 As String
Dim estrnro2 As String
Dim estrnro3 As String
Dim agrupa1 As Boolean
Dim agrupa2 As Boolean
Dim agrupa3 As Boolean
Dim pgtinro As String
Dim listaProc As String
Dim orden As String
Dim filtro_join
Dim l_arrgpanro
Dim l_i
Dim titulo As String
Dim detallado As String
Dim terape
Dim terape2
Dim ternom
Dim ternom2
Dim empleg
Dim autoriza
Dim cantReg As Long
Dim cantEmpl As Long
Dim CEmpleadosAProc As Double
Dim IncPorc As Double
Dim progreso As Double

'----------------------------SEPARO LOS PARAMETROS--------------------------------
Flog.writeline "Comienza a levantar los parametros"
Flog.writeline "Lista Parametros: " & parametros

listaParametros = Split(parametros, "@")


filtro = listaParametros(0)
Flog.writeline "Filtro: " & filtro
Flog.writeline ""

listaHoras = listaParametros(1)
Flog.writeline "listaHoras: " & listaHoras
Flog.writeline ""

tenro1 = listaParametros(2)
Flog.writeline "tenro1: " & tenro1
Flog.writeline ""

estrnro1 = listaParametros(3)
Flog.writeline "estrnro1: " & estrnro1
Flog.writeline ""

agrupa1 = listaParametros(4)
Flog.writeline "agrupa1: " & agrupa1
Flog.writeline ""

tenro2 = listaParametros(5)
Flog.writeline "tenro2: " & tenro2
Flog.writeline ""

estrnro2 = listaParametros(6)
Flog.writeline "estrnro2: " & estrnro2
Flog.writeline ""

agrupa2 = listaParametros(7)
Flog.writeline "agrupa2: " & agrupa2
Flog.writeline ""

tenro3 = listaParametros(8)
Flog.writeline "tenro3: " & tenro3
Flog.writeline ""

estrnro3 = listaParametros(9)
Flog.writeline "estrnro3: " & estrnro3
Flog.writeline ""

agrupa3 = listaParametros(10)
Flog.writeline "agrupa3: " & agrupa3
Flog.writeline ""

pgtinro = listaParametros(11)
Flog.writeline "pgtinro: " & pgtinro
Flog.writeline ""

listaProc = listaParametros(12)
Flog.writeline "listaProc: " & listaProc
Flog.writeline ""

detallado = listaParametros(13)
Flog.writeline "detallado: " & detallado
Flog.writeline ""

orden = listaParametros(14)
Flog.writeline "orden: " & orden
Flog.writeline ""

autoriza = listaParametros(15)
Flog.writeline "autoriza: " & autoriza
Flog.writeline ""


Flog.writeline "Se levantaron los parametros correctamente"
filtro = Replace(filtro, "v_empleado", "empleado")
orden = Replace(orden, "v_empleado", "empleado")
'---------------------------------------------------------------------------------

'--------------------------armo el titulo del reporte-----------------------------
titulo = "Bpronro: " & NroProceso & " - Periodo "

StrSql = "SELECT pgtimes, pgtianio, pgtidesabr FROM gti_per "
StrSql = StrSql & " WHERE pgtinro=" & pgtinro
rs2.Open StrSql, objConn
If Not rs2.EOF Then
    titulo = titulo & rs2!pgtidesabr
End If
rs2.Close
'-----------------------------------hasta aca-------------------------------------

'---------------BUSCO LOS EMPLEADOS QUE CUMPLEN CON EL FILTRO---------------------


filtro_join = " FROM gti_procacum "
filtro_join = filtro_join & " INNER JOIN gti_per ON gti_per.pgtinro = gti_procacum.pgtinro "
filtro_join = filtro_join & " INNER JOIN gti_cab ON gti_cab.gpanro  = gti_procacum.gpanro "
filtro_join = filtro_join & " AND gti_procacum.pgtinro = " & pgtinro

If listaProc <> "" Then
   l_arrgpanro = Split(listaProc, ",")
   filtro_join = filtro_join + " AND ( "
   For l_i = 0 To UBound(l_arrgpanro)
     If (l_i = 0) Then
       filtro_join = filtro_join + " gti_procacum.gpanro = " + l_arrgpanro(l_i)
     Else
       filtro_join = filtro_join + " OR gti_procacum.gpanro = " + l_arrgpanro(l_i)
     End If
   Next
   filtro_join = filtro_join + " ) "
End If

filtro_join = filtro_join + "  INNER JOIN empleado ON empleado.ternro = gti_cab.ternro "

'END - Genero un filtro para optimizar la consulta

If tenro3 <> "" And tenro3 <> "0" Then
    StrSql = "SELECT empleado.ternro, empleg, terape, ternom, estact1.tenro tenro1, estact1.estrnro  estrnro1, "
    StrSql = StrSql & " estact2.tenro  tenro2, estact2.estrnro  estrnro2, estact3.tenro  tenro3, estact3.estrnro  estrnro3 "
    StrSql = StrSql & filtro_join
    StrSql = StrSql & " INNER JOIN his_estructura estact1 ON empleado.ternro = estact1.ternro AND estact1.htethasta IS NULL AND estact1.tenro  = " & tenro1
    
    If estrnro1 <> "" And estrnro1 <> "0" Then
        StrSql = StrSql & " AND estact1.estrnro =" & estrnro1
    End If
    
    StrSql = StrSql & " INNER JOIN his_estructura estact2 ON empleado.ternro = estact2.ternro AND estact2.htethasta IS NULL AND estact2.tenro  = " & tenro2
    If estrnro2 <> "" And estrnro2 <> "0" Then
        StrSql = StrSql & " AND estact2.estrnro =" & estrnro2
    End If
    
    StrSql = StrSql & " INNER JOIN his_estructura estact3 ON empleado.ternro = estact3.ternro AND estact3.htethasta IS NULL AND estact3.tenro  = " & tenro3
    If estrnro3 <> "" And estrnro3 <> "0" Then
        StrSql = StrSql & " AND estact3.estrnro =" & estrnro3
    End If
    StrSql = StrSql & " WHERE " & filtro
    StrSql = StrSql & " GROUP BY empleado.ternro,empleg, terape, ternom,estact1.tenro, estact1.estrnro,estact2.tenro, estact2.estrnro,estact3.tenro, estact3.estrnro "
    StrSql = StrSql & " ORDER BY tenro1,estrnro1,tenro2,estrnro2,tenro3,estrnro3," & orden

Else
    If tenro2 <> "" And tenro2 <> "0" Then
        StrSql = " SELECT empleado.ternro,empleg, terape, ternom, estact1.tenro  tenro1, estact1.estrnro  estrnro1, "
        StrSql = StrSql & " estact2.tenro  tenro2, estact2.estrnro  estrnro2 "
        StrSql = StrSql & filtro_join
        StrSql = StrSql & " INNER JOIN his_estructura estact1 ON empleado.ternro = estact1.ternro AND estact1.htethasta IS NULL AND estact1.tenro  = " & tenro1 & " AND " & filtro
        
        If estrnro1 <> "" And estrnro1 <> "0" Then
            StrSql = StrSql & " AND estact1.estrnro =" & estrnro1
        End If
        StrSql = StrSql & " INNER JOIN his_estructura estact2 ON empleado.ternro = estact2.ternro AND estact2.htethasta IS NULL AND estact2.tenro  = " & tenro2
        
        If estrnro2 <> "" And estrnro2 <> "0" Then
            StrSql = StrSql & " AND estact2.estrnro =" & estrnro2
        End If
        StrSql = StrSql & " WHERE " & filtro
        StrSql = StrSql & " GROUP BY empleado.ternro,empleg, terape, ternom,estact1.tenro, estact1.estrnro,estact2.tenro, estact2.estrnro "
        StrSql = StrSql & " ORDER BY tenro1,estrnro1,tenro2,estrnro2, " & orden

     Else
        If tenro1 <> "" And tenro1 <> "0" Then
                StrSql = "SELECT empleado.ternro,empleg, terape, ternom, estact1.tenro  tenro1, estact1.estrnro  estrnro1 "
                StrSql = StrSql & filtro_join
                StrSql = StrSql & " INNER JOIN his_estructura estact1 ON empleado.ternro = estact1.ternro AND estact1.htethasta IS NULL AND estact1.tenro  = " & tenro1 & " AND " & filtro
                If estrnro1 <> "" And estrnro1 <> "0" Then
                    StrSql = StrSql & " AND estact1.estrnro =" & estrnro1
                End If
                StrSql = StrSql & " WHERE " & filtro
                StrSql = StrSql & " GROUP BY empleado.ternro,empleg, terape, ternom,estact1.tenro, estact1.estrnro "
                StrSql = StrSql & " ORDER BY tenro1,estrnro1," & orden

          Else
            StrSql = "SELECT empleado.ternro,empleg, terape, ternom "
            StrSql = StrSql & filtro_join
            StrSql = StrSql & " WHERE " & filtro
            StrSql = StrSql & " GROUP BY empleado.ternro,empleg, terape, ternom "
            StrSql = StrSql & " ORDER BY " & orden
          End If
    End If
End If
rs2.Open StrSql, objConn
MyBeginTrans
If Not rs2.EOF Then
    'Determino la proporcion de progreso
    progreso = 0
    CEmpleadosAProc = rs2.RecordCount
    If CEmpleadosAProc = 0 Then
        CEmpleadosAProc = 1
    End If
    IncPorc = (99 / CEmpleadosAProc)
    'si no es fin de archivo escribo en la tabla cabecera
    StrSql = "INSERT INTO rep_acumParcial_cab "
    StrSql = StrSql & " (bpronro, titulo, tenro1, estrnro1, agrupa1, tenro2, estrnro2, agrupa2, tenro3, estrnro3, agrupa3, periodo, autoriza)"
    StrSql = StrSql & " VALUES"
    StrSql = StrSql & "("
    StrSql = StrSql & NroProceso & ", "
    StrSql = StrSql & "'" & titulo & "', "
    StrSql = StrSql & tenro1 & ", "
    StrSql = StrSql & estrnro1 & ", "
    StrSql = StrSql & IIf(agrupa1, -1, 0) & ", "
    StrSql = StrSql & tenro2 & ", "
    StrSql = StrSql & estrnro2 & ", "
    StrSql = StrSql & IIf(agrupa2, -1, 0) & ", "
    StrSql = StrSql & tenro3 & ", "
    StrSql = StrSql & estrnro3 & ", "
    StrSql = StrSql & IIf(agrupa3, -1, 0) & ", "
    StrSql = StrSql & "'" & pgtinro & "', "
    StrSql = StrSql & autoriza
    StrSql = StrSql & ")"
    objConn.Execute StrSql, , adExecuteNoRecords
    Flog.writeline "Se inserto la cabecera del reporte "
    Flog.writeline ""
    
    Do While Not rs2.EOF
        'guardo el tipo estructura y la estructura de cada empleado en caso de tener
        'estructura 1
        If tenro1 <> "0" Then
            tenro1 = IIf(EsNulo(rs2!tenro1), "", rs2!tenro1)
            estrnro1 = IIf(EsNulo(rs2!estrnro1), "", rs2!estrnro1)
        End If
        
        'estructura 2
        If tenro2 <> "0" Then
            tenro2 = IIf(EsNulo(rs2!tenro2), "", rs2!tenro2)
            estrnro2 = IIf(EsNulo(rs2!estrnro2), "", rs2!estrnro2)
        End If
        
        'estructura 3
        If tenro3 <> "0" Then
            tenro3 = IIf(EsNulo(rs2!tenro3), "", rs2!tenro3)
            estrnro3 = IIf(EsNulo(rs2!estrnro3), "", rs2!estrnro3)
        End If
        'hasta aca
    
        StrSqlAux = "SELECT ternro, thdesc, tiphora.thnro, dgticant, gti_procacum.gpanro,gpadesabr,pgtidesabr "
        StrSqlAux = StrSqlAux & " FROM gti_procacum "
        StrSqlAux = StrSqlAux & " INNER JOIN gti_per ON gti_per.pgtinro = gti_procacum.pgtinro "
        StrSqlAux = StrSqlAux & " INNER JOIN gti_cab ON gti_cab.gpanro  = gti_procacum.gpanro "
        StrSqlAux = StrSqlAux & " INNER JOIN gti_det ON gti_det.cgtinro = gti_cab.cgtinro "
        StrSqlAux = StrSqlAux & " INNER JOIN tiphora ON tiphora.thnro   = gti_det.thnro "
        StrSqlAux = StrSqlAux & " WHERE gti_procacum.pgtinro = " & pgtinro
        StrSqlAux = StrSqlAux & " AND gti_cab.ternro = " & rs2!Ternro
    
        If listaProc <> "" Then
           l_arrgpanro = Split(listaProc, ",")
           StrSqlAux = StrSqlAux + " AND ( "
           For l_i = 0 To UBound(l_arrgpanro)
             If (l_i = 0) Then
               StrSqlAux = StrSqlAux + " gti_procacum.gpanro = " + l_arrgpanro(l_i)
             Else
               StrSqlAux = StrSqlAux + " OR gti_procacum.gpanro = " + l_arrgpanro(l_i)
             End If
           Next
           StrSqlAux = StrSqlAux + " ) "
        End If
        StrSqlAux = StrSqlAux & " ORDER BY gti_procacum.gpanro, gti_cab.ternro, thdesc"
        rs3.Open StrSqlAux, objConn
        If Not rs3.EOF Then
            'busco los datos del empleado
            sqlaux2 = " SELECT * FROM empleado "
            sqlaux2 = sqlaux2 & " WHERE ternro=" & rs2!Ternro
            rs4.Open sqlaux2, objConn
            If Not rs2.EOF Then
                terape = IIf(EsNulo(rs4!terape), "", rs4!terape)
                terape2 = IIf(EsNulo(rs4!terape2), "", rs4!terape2)
                ternom = IIf(EsNulo(rs4!ternom), "", rs4!ternom)
                ternom2 = IIf(EsNulo(rs4!ternom2), "", rs4!ternom2)
                empleg = IIf(EsNulo(rs4!empleg), "", rs4!empleg)
            End If
            rs4.Close
            'hasta aca
            Do While Not rs3.EOF
                cantReg = cantReg + 1
                StrSql = " INSERT INTO rep_acumParcial_det "
                StrSql = StrSql & "( bpronro, ternro, empleg, terape, terape2, ternom, ternom2, tenro1, estrnro1, tenro2, estrnro2, tenro3, estrnro3, thdesc, thnro, dgticant,gpanro,gpadesabr,pgtidesabr,pgtinro)"
                StrSql = StrSql & " VALUES "
                StrSql = StrSql & "("
                StrSql = StrSql & NroProceso & ", "
                StrSql = StrSql & rs2!Ternro & ", "
                StrSql = StrSql & empleg & ", "
                StrSql = StrSql & "'" & terape & "' , "
                StrSql = StrSql & "'" & terape2 & "' , "
                StrSql = StrSql & "'" & ternom & "' , "
                StrSql = StrSql & "'" & ternom2 & "' , "
                StrSql = StrSql & tenro1 & ", "
                StrSql = StrSql & estrnro1 & ", "
                StrSql = StrSql & tenro2 & ", "
                StrSql = StrSql & estrnro2 & ", "
                StrSql = StrSql & tenro3 & ", "
                StrSql = StrSql & estrnro3 & ", "
                StrSql = StrSql & "'" & rs3!thdesc & "' , "
                StrSql = StrSql & rs3!thnro & " , "
                StrSql = StrSql & rs3!dgticant & ", "
                StrSql = StrSql & rs3!gpanro & ", "
                StrSql = StrSql & "'" & rs3!gpadesabr & "' , "
                StrSql = StrSql & "'" & rs3!pgtidesabr & "' ,  "
                StrSql = StrSql & pgtinro
                StrSql = StrSql & ")"
                objConn.Execute StrSql, , adExecuteNoRecords
                Flog.writeline "Se inserto el detalle para el empleado: " & empleg
                Flog.writeline ""
            rs3.MoveNext
            Loop
        Else
            Flog.writeline "No hay datos de horas para el tercero: " & rs2!Ternro
            Flog.writeline ""
        End If
        rs3.Close
        
        'Incremento el progreso para que el servidor de aplicaciones no vea a este proceso
        'como colgado
        progreso = progreso + IncPorc
        Flog.writeline Espacios(Tabulador * 0) & "Progreso = " & CLng(progreso) & " (Incremento = " & IncPorc & ")"
        StrSql = "UPDATE batch_proceso SET bprcprogreso = " & CLng(progreso) & " WHERE bpronro = " & NroProceso
        objConnProgreso.Execute StrSql, , adExecuteNoRecords
        Flog.writeline Espacios(Tabulador * 0) & "Progreso actualizado"
    
    rs2.MoveNext
    Loop
Else
    Flog.writeline "No hay empleados que cumplan con el filtro"
End If
rs2.Close
'--------------------------------HASTA ACA----------------------------------------

'si no habia datos de los empleados para insertar borro la cabecera
If cantReg = 0 Then
    MyRollbackTrans
    Flog.writeline "No hay datos para los empleados del filtro."
Else
    MyCommitTrans
    Flog.writeline "Se insertaron los datos correctamente."
End If
'hasta aca


End Sub
