Attribute VB_Name = "mdlEvoNominaPla"
Option Explicit

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Const Version = "1.00" 'Version inicial
Const FechaVersion = "21/07/2014"
'CAS-24476 - PLA - Nuevo Reporte de Evolucion de Nomina - LED
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Dim fs
Dim Flog
Dim FDesde As Date
Dim FHasta As Date
Global Fecha_Inicio As Date
Global CEmpleadosAProc As Long
Global IncPorc As Double
Global Progreso As Double
Global TiempoInicialProceso As Long
Global totalEmpleados As Long
Global TiempoAcumulado As Long
Global cantRegistros As Long
Dim l_iduser
Dim l_estrnro1
Dim l_estrnro2
Dim l_estrnro3




Sub Main()

Dim Archivo As String
Dim pos As Integer
Dim strcmdLine  As String
Dim rs As New ADODB.Recordset

Dim NroProceso As Long
Dim NroReporte As Long
Dim StrParametros As String

Dim rs_Batch_Proceso As New ADODB.Recordset
Dim rs_His_Batch_Proceso As New ADODB.Recordset

Dim PID As String
Dim ArrParametros


    strcmdLine = Command()
    ArrParametros = Split(strcmdLine, " ", -1)
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
            If IsNumeric(strcmdLine) Then
                NroProceso = strcmdLine
            Else
                Exit Sub
            End If
        End If
    End If
    
    
    ' carga las configuraciones basicas, formato de fecha, string de conexion,
    ' tipo de BD y ubicacion del archivo de log
    Call CargarConfiguracionesBasicas

    'Creo el archivo de texto del desglose
    Archivo = PathFLog & "EvoNomina_" & CStr(NroProceso) & ".log"

    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Archivo, True)


    'Abro la conexion
    On Error Resume Next
    OpenConnection strconexion, objConn
    If Err.Number <> 0 Or Error_Encrypt Then
        Flog.writeline Espacios(Tabulador * 0) & "Problemas en la conexion"
        Exit Sub
    End If
    On Error Resume Next
    OpenConnection strconexion, objConnProgreso
    If Err.Number <> 0 Or Error_Encrypt Then
        Flog.writeline Espacios(Tabulador * 0) & "Problemas en la conexion"
        Exit Sub
    End If

    On Error GoTo ce


    'Obtengo el Process ID
    PID = GetCurrentProcessId
    Flog.writeline "-------------------------------------------------"
    Flog.writeline "Version                  : " & Version
    Flog.writeline "Fecha Ultima Modificacion: " & FechaVersion
    Flog.writeline "PID                      : " & PID
    Flog.writeline "-------------------------------------------------"
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 0) & "Inicio :" & Now
    
    'FGZ - 05/08/2009 --------- Control de versiones ------
    Version_Valida = ValidarV(Version, 1, TipoBD)
    If Not Version_Valida Then
        'Actualizo el progreso
        StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcprogreso = 0, bprcestado = 'Error de Version', bprcpid = " & PID & " WHERE bpronro = " & NroProceso
        objConnProgreso.Execute StrSql, , adExecuteNoRecords
        Flog.writeline
        GoTo Final
    End If
    'FGZ - 05/08/2009 --------- Control de versiones ------

    'Cambio el estado del proceso a Procesando
    StrSql = "UPDATE batch_proceso SET bprcpid = " & PID & ", bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcestado = 'Procesando', bprcprogreso = 0 WHERE bpronro = " & NroProceso
    objConn.Execute StrSql, , adExecuteNoRecords

    Flog.writeline "Levanta Proceso y Setea Parámetros:  " & " " & Now

    'levanto los parametros del proceso
    StrParametros = ""

    StrSql = "SELECT bprcfecdesde,bprcfechasta,bprcparam,iduser FROM batch_proceso WHERE bpronro = " & NroProceso
    OpenRecordset StrSql, rs
    If Not rs.EOF Then
        Flog.writeline "Inicio de Reporte evolucion de nomina: " & Now
        Call Rep_evolucion_nomina(rs!bprcparam, NroProceso)
    Else
        Exit Sub
    End If
        
    ' poner el bprcestado en procesado
    StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado', bprcprogreso = 100 WHERE bpronro = " & NroProceso
    objConn.Execute StrSql, , adExecuteNoRecords

    If objConn.State = adStateOpen Then objConn.Close
    If objConnProgreso.State = adStateOpen Then objConnProgreso.Close
    
    Set rs_Batch_Proceso = Nothing
    Set rs_His_Batch_Proceso = Nothing

Final:
    Flog.writeline Espacios(Tabulador * 0) & "Fin de Reporte de evolucion de nomina: " & " " & Now
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 0) & "---------------------------------------------------------------------------------"
    Flog.writeline Espacios(Tabulador * 0) & "---------------------------------------------------------------------------------"
    Flog.Close
Exit Sub

ce:
    Flog.writeline Espacios(Tabulador * 0) & "Reporte abortado por Error:" & " " & Now
    Flog.writeline Espacios(Tabulador * 0) & "Reporte abortado por :" & Err.Description
    Flog.writeline Espacios(Tabulador * 0) & "Ultimo SQL " & StrSql
    StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Error' WHERE bpronro = " & NroProceso
    objConn.Execute StrSql, , adExecuteNoRecords
End Sub
Private Sub Rep_evolucion_nomina(parametros As String, NroProceso As Long)

Dim rs As New ADODB.Recordset
Dim legdesde As String
Dim leghasta As String
Dim estado As String
Dim Tenro As String
Dim estrnro As String
Dim fecdesde As String
Dim fechasta As String
Dim ArrParametros
Dim historicoDesc As String
Dim mes As Long
Dim empest As String
Dim cantEstructuras As String
Dim cantMeses As Long
Dim empEstrnro As Long
Dim empresa As String
'======================================================================================================
'SE VALIDA Y LEVANTAN PARAMETROS.
'======================================================================================================
        
    Flog.writeline ""
    Flog.writeline "Parametros:" & parametros
    Flog.writeline ""
    
    ' parametros(0) --> legdesde
    ' parametros(1) --> leghasta
    ' parametros(2) --> estado
    ' parametros(3) --> tenro1
    ' parametros(4) --> estrnro1
    ' parametros(5) --> fecha desde
    ' parametros(6) --> fecha hasta

    
    ArrParametros = Split(parametros, "@")
    
    legdesde = ArrParametros(0)
    leghasta = ArrParametros(1)
    estado = ArrParametros(2)
    Tenro = ArrParametros(3)
    estrnro = ArrParametros(4)
    fecdesde = ArrParametros(5)
    fechasta = ArrParametros(6)
    empEstrnro = ArrParametros(7)
    
    Flog.writeline "Parametros Obtenidos."
    
    If estado = -1 Then
        empest = " AND empleado.empest = -1 "
    End If
    
    If estado = 0 Then
        empest = " AND empleado.empest = 0 "
    End If
    
    StrSql = " SELECT tedabr FROM tipoestructura WHERE tenro = " & Tenro
    OpenRecordset StrSql, rs
    If Not rs.EOF Then
        historicoDesc = NroProceso & " - " & rs!tedabr & " - Evolucion Nomina del " & fecdesde & " al " & fechasta
    Else
        Flog.writeline "No existe el tipo de estructura: " & Tenro
        Exit Sub
    End If
    
    fecdesde = DateSerial(Year(fecdesde), Month(fecdesde) + 1, 0)
    fechasta = DateSerial(Year(fechasta), Month(fechasta) + 1, 0)
    cantMeses = DateDiff("m", fecdesde, fechasta)
    cantMeses = Year(fechasta) * 12 + Month(fechasta) - (Year(fecdesde) * 12 + Month(fecdesde)) + 1
    Flog.writeline "Fechas configurada"
    
    
    Flog.writeline "Calculo de progreso"
    Progreso = 0
    IncPorc = (100 / cantMeses)
    TiempoInicialProceso = GetTickCount
        
        
        
    If CLng(estrnro) <> -1 Then
        estrnro = "0," & estrnro
    Else
        StrSql = " SELECT estrnro FROM estructura WHERE tenro = " & Tenro & " ORDER BY estrnro "
        OpenRecordset StrSql, rs
        estrnro = "0"
        cantEstructuras = rs.RecordCount
        Do While Not rs.EOF
            estrnro = estrnro & "," & rs!estrnro
            rs.MoveNext
        Loop
    End If
    Flog.writeline "Lista de estructura obtenida para el tenro: " & Tenro
    
    'obtengo el nombre de la empresa
    StrSql = " SELECT estrdabr FROM estructura WHERE estrnro = " & empEstrnro
    OpenRecordset StrSql, rs
    If Not rs.EOF Then
        empresa = rs!estrdabr
    End If
    Flog.writeline "Nombre de la empresa obtenido: " & empresa
    
    'inserto la cabecera del reporte
    StrSql = " INSERT INTO rep_evo_nomina (bpronro, descripcion, empresa, fecgen, fecdesde, fechasta) VALUES " & _
             " (" & NroProceso & ",'" & historicoDesc & "',' " & empresa & "'," & ConvFecha(Date) & "," & ConvFecha(fecdesde) & "," & ConvFecha(fechasta) & ") "
    objConn.Execute StrSql, , adExecuteNoRecords
    Flog.writeline "Cabecera del reporte insertada"
    
    Do Until CDate(fechasta) < CDate(fecdesde)
        Progreso = Progreso + IncPorc
        
        Flog.writeline "Analizando estructuras para la fecha: " & fechasta
        StrSql = " SELECT count(his_estructura.ternro) cant, his_estructura.estrnro, estructura.estrdabr FROM empleado " & _
                 " INNER JOIN his_estructura emp ON emp.ternro = empleado .ternro AND emp.tenro = 10 " & _
                 " INNER JOIN his_estructura ON his_estructura.ternro = empleado .ternro AND his_estructura.tenro = " & Tenro & _
                 " INNER JOIN tipoestructura ON tipoestructura.tenro = his_estructura.tenro " & _
                 " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro " & _
                 " WHERE (his_estructura.htetdesde <= " & ConvFecha(fechasta) & " AND (his_estructura.htethasta >= " & ConvFecha(fechasta) & " OR his_estructura.htethasta is null)) " & _
                 " AND his_estructura.estrnro in (" & estrnro & ") " & _
                 " AND (emp.htetdesde <= " & ConvFecha(fechasta) & " AND (emp.htethasta >= " & ConvFecha(fechasta) & " OR emp.htethasta is null)) AND emp.estrnro = " & empEstrnro & _
                 " AND (empleado.empleg >= " & legdesde & " AND empleado.empleg <= " & leghasta & ") " & empest & _
                 " group by his_estructura.estrnro, estructura.estrdabr, tipoestructura.tedabr " & _
                 " order by estructura.estrdabr "
        OpenRecordset StrSql, rs
                
        Do While Not rs.EOF
            StrSql = " INSERT INTO rep_evo_nomina_det (bpronro, tenro, estrnro, mes, cant) VALUES " & _
                     " (" & NroProceso & "," & Tenro & "," & rs!estrnro & "," & Month(fechasta) & "," & rs!Cant & ")"
            objConn.Execute StrSql, , adExecuteNoRecords
            Flog.writeline "Dato insertado para la estructura: " & rs!estrnro & " y mes: " & Month(fechasta)
            rs.MoveNext
        Loop
        
        
        'Actualizo el progreso
         TiempoAcumulado = GetTickCount
         StrSql = "UPDATE batch_proceso SET bprcprogreso = " & CDbl(Progreso)
         StrSql = StrSql & ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'"
         StrSql = StrSql & ", bprcempleados ='" & CStr(CEmpleadosAProc) & "' WHERE bpronro = " & NroProceso
         objConnProgreso.Execute StrSql, , adExecuteNoRecords

        
        fechasta = DateAdd("m", -1, fechasta)
        fechasta = DateSerial(Year(fechasta), Month(fechasta) + 1, 0)
        
    Loop


Flog.writeline ""
Flog.writeline "Proceso finalizado."


If rs.State = adStateOpen Then rs.Close

Set rs = Nothing

End Sub






Function empleadosEstructura(ByVal estrnro1 As String, ByVal estrnro2 As String, ByVal fecdesde As String, ByVal fechasta As String)
Dim rsEst As New ADODB.Recordset

    StrSql = " SELECT count(distinct empleado.ternro) cant FROM empleado " & _
            " INNER JOIN his_estructura he1 ON he1.ternro = empleado.ternro AND he1.estrnro = " & estrnro1 & _
            " INNER JOIN his_estructura he2 ON he2.ternro = empleado.ternro AND he2.estrnro = " & estrnro2 & _
            " WHERE ((he1.htetdesde <= " & ConvFecha(fecdesde) & " AND (he1.htethasta is null or he1.htethasta >= " & ConvFecha(fechasta) & _
            " or he1.htethasta >= " & ConvFecha(fecdesde) & ")) OR(he1.htetdesde >= " & ConvFecha(fecdesde) & " AND (he1.htetdesde <= " & ConvFecha(fechasta) & "))) " & _
            " AND ((he2.htetdesde <= " & ConvFecha(fecdesde) & " AND (he2.htethasta is null or he2.htethasta >= " & ConvFecha(fechasta) & _
            " or he2.htethasta >= " & ConvFecha(fecdesde) & ")) OR(he2.htetdesde >= " & ConvFecha(fecdesde) & " AND (he2.htetdesde <= " & ConvFecha(fechasta) & "))) " & _
            " AND empest = -1 "
    OpenRecordset StrSql, rsEst
    If Not rsEst.EOF Then
        If Not IsNull(rsEst!Cant) Then
            empleadosEstructura = rsEst!Cant
        Else
            empleadosEstructura = 0
        End If
    Else
        empleadosEstructura = 0
    End If


If rsEst.State = adStateOpen Then rsEst.Close
Set rsEst = Nothing

End Function

