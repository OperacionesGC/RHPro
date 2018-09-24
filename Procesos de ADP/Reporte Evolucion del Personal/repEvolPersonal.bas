Attribute VB_Name = "repEvolPersonal"
'Global Const Version = "1.00" ' Carmen Quintero
'Global Const FechaModificacion = "13/08/2012"
'Global Const UltimaModificacion = "" 'Version Inicial

Global Const Version = "1.01" ' Carmen Quintero
Global Const FechaModificacion = "30/08/2012"
Global Const UltimaModificacion = "" 'Carmen Quintero (16166) Se modificó el calculo de la dotacion mensual.


'--------------------------------------------------------------
'--------------------------------------------------------------
Option Explicit

Dim fs, f
'Global Flog

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
Global EmpErrores As Boolean

Global Pagina As Long
Global tipoModelo As Integer
Global arrTipoConc(1000) As Integer
Global tituloReporte As String

Global tenro1 As Integer
Global estrnro1 As Integer
Global tenro2 As Integer
Global estrnro2 As Integer
Global tenro3 As Integer
Global estrnro3 As Integer
Global fecEstr As String
Global fechadesde As String
Global fechahasta As String
Global agencia As Integer

Global empresa As String
Global Empnro As Long
Global Empnroestr As Long
Global Centcostnroestr As Long
Global emprTer As Long
Global emprDire As String
Global emprCuit

Global IdUser As String
Global Fecha As Date
Global Hora As String

Global listapronro       'Lista de procesos

Global totalEmpleados
Global cantRegistros

Global incluyeAgencia As Integer
Global NroAcDiasTrabajados As Long


Global CantEmpGrabados As Long 'Cantidad de empleados grabados

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
Dim objRs3 As New ADODB.Recordset

Dim historico As Boolean
'Dim param
Dim proNro As Long
Dim ternro  As Long
Dim arrpronro
Dim rsEmpl As New ADODB.Recordset
Dim rsAge As New ADODB.Recordset
Dim rsEmpresas As New ADODB.Recordset
Dim rsPeriodo As New ADODB.Recordset
'Dim acunroSueldo
Dim I
Dim PID As String

Dim parametros As String
Dim ArrParametros
Dim strTempo As String
Dim orden As String

    
Dim arrpliqnro
Dim listapliqnro
Dim pliqNro As Long
Dim pliqMes As Long
Dim pliqAnio As Long
Dim rsConsult2 As New ADODB.Recordset

    
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
    
    TiempoInicialProceso = GetTickCount
    
    Nombre_Arch = PathFLog & "ReporteEvolPersonal" & "-" & NroProceso & ".log"
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
    
    Flog.writeline "Inicio Proceso: " & Now
    Flog.writeline "Cambio el estado del proceso a Procesando"
    
    'OpenConnection strconexion, objConn
    On Error Resume Next
    OpenConnection strconexion, objConn
    If Err.Number <> 0 Or Error_Encrypt Then
        Flog.writeline Espacios(Tabulador * 0) & "Problemas en la conexion"
        Exit Sub
    End If
    
    HuboErrores = False
    
    On Error GoTo CE
    
    'Cambio el estado del proceso a Procesando
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcestado = 'Procesando', bprcpid = " & PID & " WHERE bpronro = " & NroProceso
    objConn.Execute StrSql, , adExecuteNoRecords
    
    Flog.writeline "Obtengo los datos del proceso"
    
    TiempoAcumulado = GetTickCount
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE bpronro = " & NroProceso
    OpenRecordset StrSql, objRs
    
    If Not objRs.EOF Then
       
       'Obtengo los parametros del proceso
       IdUser = objRs!IdUser
       Fecha = objRs!bprcfecha
       Hora = objRs!bprchora
       parametros = objRs!bprcparam
       Flog.writeline " parametros del proceso --> " & parametros
       ArrParametros = Split(parametros, "@")
       Flog.writeline " limite del array --> " & UBound(ArrParametros)
       
       
       'Obtengo el tipo de estructura 1 si se configuró
       If CLng(ArrParametros(1)) <> 0 Then
            tenro1 = CLng(ArrParametros(1))
            Flog.writeline "Se selecciono el parametro Tipo de Estructura 1. " & ArrParametros(1)
       End If
       
       'Obtengo la estructura 1 si se configuró
       If CLng(ArrParametros(2)) <> 0 Then
            estrnro1 = CLng(ArrParametros(2))
            Flog.writeline "Se selecciono el parametro Estructura 1. " & ArrParametros(2)
       End If
       
       'Obtengo el tipo de estructura 2 si se configuró
       If CLng(ArrParametros(3)) <> 0 Then
            tenro2 = CLng(ArrParametros(3))
            Flog.writeline "Se selecciono el parametro Tipo de Estructura 2. " & ArrParametros(3)
       End If
       
       'Obtengo la estructura 2 si se configuró
       If CLng(ArrParametros(4)) <> 0 Then
            estrnro2 = CLng(ArrParametros(4))
            Flog.writeline "Se selecciono el parametro Estructura 2. " & ArrParametros(4)
       End If
       
       'Obtengo el tipo de estructura 3 si se configuró
       If CLng(ArrParametros(5)) <> 0 Then
            tenro3 = CLng(ArrParametros(5))
            Flog.writeline "Se selecciono el parametro Tipo de Estructura 3. " & ArrParametros(5)
       End If
       
       'Obtengo la estructura 3 si se configuró
       If CLng(ArrParametros(6)) <> 0 Then
            estrnro3 = CLng(ArrParametros(6))
            Flog.writeline "Se selecciono el parametro Estructura 3. " & ArrParametros(6)
       End If
       
       'Obtengo la fecha desde
       fechadesde = ArrParametros(7)
       Flog.writeline "Se selecciono el parametro fecha desde. " & ArrParametros(7)
       If Len(fechadesde) = 0 Then
            Flog.writeline "No Se selecciono el parametro fecha desde. "
            HuboErrores = True
       End If

       'Obtengo la fecha hasta
       fechahasta = ArrParametros(8)
       Flog.writeline "Se selecciono el parametro fecha hasta. " & ArrParametros(8)
       If Len(fechahasta) = 0 Then
            Flog.writeline "No Se selecciono el parametro fecha hasta. "
            HuboErrores = True
       End If
       
       'Obtengo la agencia
       agencia = ArrParametros(9)
       Flog.writeline "Se selecciono el parametro agencia. " & ArrParametros(9)
       If CLng(agencia) = 0 Then
            Flog.writeline "No Se selecciono el parametro agencia. "
       End If

       
       'Obtengo el orden
       orden = ArrParametros(10)
       Flog.writeline "Se selecciono el parametro orden. " & ArrParametros(10)
       
       tituloReporte = ArrParametros(11)
       Flog.writeline "Se selecciono el parametro titulo del reporte. " & ArrParametros(11)
              
       'EMPIEZA EL PROCESO
       Flog.writeline "Generando el reporte"
                    
                  
       'Obtengo los empleados sobre los que tengo que generar el reporte
       'CargarEmpleados(ByVal NroProc, ByRef rsEmpl As ADODB.Recordset, ByVal empresa As Long)
       CargarEmpleados NroProceso, rsEmpl, 0
       If Not rsEmpl.EOF Then
            'cantRegistros = rsEmpl.RecordCount
            Flog.writeline "Cantidad de empleados a procesar: " & cantRegistros
            CantEmpGrabados = 0 'Cantidad de empleados Guardados
       Else
            Flog.writeline "No hay empleados para el filtro seleccionado."
            Exit Sub
       End If
    
       'Actualizo Barch Proceso
       StrSql = "UPDATE batch_proceso SET bprcprogreso = 0 " & _
                ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'" & _
                ", bprcempleados ='" & CStr(rsEmpl.RecordCount) & "' WHERE bpronro = " & NroProceso
    
       objConn.Execute StrSql, , adExecuteNoRecords
 
       'Verifico que batch_empleado tenga registros
       If Not rsEmpl.EOF Then
            EmpErrores = False
            ternro = rsEmpl!ternro
            Flog.writeline ""
            Flog.writeline "Generando datos de los empleados "
                    
            Call ReporteEvolPersonal
                                            
            'Actualizo el estado del proceso
            TiempoAcumulado = GetTickCount
            
            'Resto uno a la cantidad de registros
            cantRegistros = rsEmpl.RecordCount
            
            'Actualizo
            StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Fix(((totalEmpleados - cantRegistros) * 100) / totalEmpleados) & _
                     ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'" & _
                     ", bprcempleados ='" & CStr(cantRegistros) & "' WHERE bpronro = " & NroProceso
                    
            objConn.Execute StrSql, , adExecuteNoRecords
            
            'Borro batch empleado
            '****************************************************************
            StrSql = "DELETE  FROM batch_empleado "
            StrSql = StrSql & " WHERE bpronro = " & NroProceso
            StrSql = StrSql & " AND ternro = " & ternro
            objConn.Execute StrSql, , adExecuteNoRecords
            
       End If
       rsEmpl.Close
       Set rsEmpl = Nothing
       
       objRs.Close
       Set objRs = Nothing
    
    Else
        objRs.Close
        Set objRs = Nothing
        
        objConn.Close
        Set objConn = Nothing
        
        Exit Sub
    End If
       
    
    'Actualizo el estado del proceso
    If Not HuboErrores Then
       StrSql = "UPDATE batch_proceso SET  bprcprogreso = 100, bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado' WHERE bpronro = " & NroProceso
    Else
       StrSql = "UPDATE batch_proceso SET  bprcprogreso = 100, bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Incompleto' WHERE bpronro = " & NroProceso
    End If
    
    objConn.Execute StrSql, , adExecuteNoRecords
    Flog.writeline
    Flog.writeline "************************************************************"
    Flog.writeline "Fin :" & Now
    Flog.writeline "Cantidad de empleados guardados en el reporte: " & cantRegistros
    Flog.Close

    Exit Sub
    
CE:
    HuboErrores = True
    Flog.writeline " Error: " & Err.Description & Now
    Flog.writeline " Ultimo SQL: " & StrSql
End Sub


Function controlNull(Str)
  If Trim(Str) = "" Then
     controlNull = "null"
  Else
     controlNull = "'" & Str & "'"
  End If
End Function

Function CantidadDias(Fecha)
    Dim proxmes, finmes
    proxmes = DateAdd("m", 1, CDate(Fecha))
    finmes = proxmes - DatePart("d", proxmes)
    CantidadDias = DatePart("d", finmes)
End Function

Function EmplExistentes(l_desde, l_hasta)
Dim rs2 As New ADODB.Recordset
Dim fil_agen As String


On Error GoTo ME_Empleados

    fil_agen = "" ' cuando queremos todos los empleados

    If agencia = -1 Then
            fil_agen = " AND empleado.ternro not in (SELECT ternro from his_estructura agencia "
            fil_agen = fil_agen & " WHERE agencia.tenro=28 "
            fil_agen = fil_agen & " AND (agencia.htetdesde <= " & ConvFecha(l_hasta) & " "
            fil_agen = fil_agen & " AND (agencia.htethasta IS NULL OR agencia.htethasta >= " & ConvFecha(l_hasta) & ")) )"
    Else
        If agencia = -2 Then
            fil_agen = " AND empleado.ternro in (SELECT ternro from his_estructura agencia "
            fil_agen = fil_agen & " WHERE agencia.tenro=28 "
            fil_agen = fil_agen & " AND (agencia.htetdesde <= " & ConvFecha(l_hasta) & " "
            fil_agen = fil_agen & " AND (agencia.htethasta IS NULL OR agencia.htethasta >= " & ConvFecha(l_hasta) & ")) )"
        Else
            If agencia <> 0 Then 'este caso se da cuando selecionamos una agencia determinada
                fil_agen = " AND empleado.ternro in (SELECT ternro from his_estructura agencia "
                fil_agen = fil_agen & " WHERE agencia.tenro=28 AND agencia.estrnro=" & agencia & ""
                fil_agen = fil_agen & " AND (agencia.htetdesde <= " & ConvFecha(l_hasta) & " "
                fil_agen = fil_agen & " AND (agencia.htethasta IS NULL OR agencia.htethasta >= " & ConvFecha(l_hasta) & ")) )"
            End If
        End If
    End If

    '----------------------------------------------------------------------------------------
    ' Calculo los ingresos que tuvo el tipo de estructura
    '----------------------------------------------------------------------------------------
    
    If tenro3 <> 0 Then ' esto ocurre solo cuando se seleccionan los tres niveles
        StrSql = "SELECT COUNT (DISTINCT empleado.ternro) cantexisten "
        StrSql = StrSql & " FROM empleado "
        StrSql = StrSql & " INNER JOIN tercero ON tercero.ternro = empleado.ternro "
        StrSql = StrSql & " INNER JOIN his_estructura estact1 ON empleado.ternro = estact1.ternro  AND estact1.tenro  = " & tenro1 & ""
        StrSql = StrSql & " AND (estact1.htetdesde<=" & ConvFecha(l_hasta) & " AND (estact1.htethasta IS NULL OR estact1.htethasta>=" & ConvFecha(l_hasta) & "))"

        If estrnro1 <> 0 Then 'cuando se le asigna un valor al nivel 1
            StrSql = StrSql & " AND estact1.estrnro =" & estrnro1
        End If
        
        StrSql = StrSql & " INNER JOIN his_estructura estact2 ON empleado.ternro = estact2.ternro  AND estact2.tenro  = " & tenro2 & ""
        StrSql = StrSql & " AND (estact2.htetdesde<=" & ConvFecha(l_hasta) & " AND (estact2.htethasta IS NULL OR estact2.htethasta>=" & ConvFecha(l_hasta) & "))"
        
        If estrnro2 <> 0 Then 'cuando se le asigna un valor al nivel 2
            StrSql = StrSql & " AND estact2.estrnro =" & estrnro2
        End If
        
        StrSql = StrSql & " INNER JOIN his_estructura estact3 ON empleado.ternro = estact3.ternro  AND estact3.tenro  = " & tenro3 & ""
        StrSql = StrSql & " AND (estact3.htetdesde<=" & ConvFecha(l_hasta) & " AND (estact3.htethasta IS NULL OR estact3.htethasta>=" & ConvFecha(l_hasta) & "))"
        
        If estrnro3 <> 0 Then 'cuando se le asigna un valor al nivel 3
            StrSql = StrSql & " AND estact3.estrnro =" & estrnro3
        End If
        
        StrSql = StrSql & " WHERE " & fil_agen & ""
    
    Else
        If tenro2 <> 0 Then ' ocurre cuando se selecciono hasta el segundo nivel
            StrSql = "SELECT COUNT (DISTINCT empleado.ternro) cantexisten "
            StrSql = StrSql & " FROM empleado "
            StrSql = StrSql & " INNER JOIN tercero ON tercero.ternro = empleado.ternro "
            StrSql = StrSql & " INNER JOIN his_estructura estact1 ON empleado.ternro = estact1.ternro  AND estact1.tenro  = " & tenro1 & ""
            StrSql = StrSql & " AND (estact1.htetdesde<=" & ConvFecha(l_hasta) & " AND (estact1.htethasta IS NULL OR estact1.htethasta>=" & ConvFecha(l_hasta) & "))"

            If estrnro1 <> 0 Then
                StrSql = StrSql & " AND estact1.estrnro =" & estrnro1
            End If
            
            StrSql = StrSql & " INNER JOIN his_estructura estact2 ON empleado.ternro = estact2.ternro  AND estact2.tenro  = " & tenro2 & ""
            StrSql = StrSql & " AND (estact2.htetdesde<=" & ConvFecha(l_hasta) & " AND (estact2.htethasta IS NULL OR estact2.htethasta>=" & ConvFecha(l_hasta) & "))"
            
            If estrnro2 <> 0 Then
                StrSql = StrSql & " AND estact2.estrnro =" & estrnro2
            End If
            
            StrSql = StrSql & " WHERE " & fil_agen & ""
        Else
            If tenro1 <> 0 Then ' Cuando solo seleccionamos el primer nivel
                 StrSql = "SELECT COUNT (DISTINCT empleado.ternro) cantexisten "
                 StrSql = StrSql & " FROM empleado "
                 StrSql = StrSql & " INNER JOIN tercero ON tercero.ternro = empleado.ternro "
                 StrSql = StrSql & " INNER JOIN his_estructura estact1 ON empleado.ternro = estact1.ternro  AND estact1.tenro  = " & tenro1 & ""
                 StrSql = StrSql & " AND (estact1.htetdesde<=" & ConvFecha(l_hasta) & " AND (estact1.htethasta IS NULL OR estact1.htethasta>=" & ConvFecha(l_hasta) & "))"
                    
                If estrnro1 <> 0 Then
                    StrSql = StrSql & " AND estact1.estrnro =" & estrnro1
                End If
                
                StrSql = StrSql & " WHERE " & fil_agen & ""
                
            Else ' cuando no hay nivel de estructura seleccionado
                 StrSql = "SELECT COUNT (DISTINCT empleado.ternro) cantexisten "
                 StrSql = StrSql & " FROM empleado "
                 StrSql = StrSql & " INNER JOIN tercero ON tercero.ternro = empleado.ternro "
                 StrSql = StrSql & " INNER JOIN his_estructura estact1 ON empleado.ternro = estact1.ternro "
                 StrSql = StrSql & " WHERE (estact1.htetdesde<=" & ConvFecha(l_hasta) & " AND (estact1.htethasta IS NULL OR estact1.htethasta>=" & ConvFecha(l_hasta) & "))"
                 StrSql = StrSql & " " & fil_agen & ""
            End If
        End If
    End If
    
    'Flog.writeline " query Exitentes: " & StrSql
    'Flog.writeline " query Exitentes Tiempo: " & Timer
    
    OpenRecordset StrSql, rs2
    If Not rs2.EOF Then
        EmplExistentes = rs2!cantexisten
    Else
        EmplExistentes = 0
    End If
    rs2.Close
    

Exit Function

ME_Empleados:
    Flog.writeline "    Error: " & Err.Description
    Flog.writeline "    SQL Ejecutado: " & StrSql
    Flog.writeline "  "
    
End Function

Function EmplBajas(l_desde, l_hasta)
Dim rs2 As New ADODB.Recordset
Dim fil_agen As String


On Error GoTo ME_fases

    fil_agen = "" ' cuando queremos todos los empleados
    If agencia = -1 Then
            fil_agen = " AND empleado.ternro not in (SELECT ternro from his_estructura agencia "
            fil_agen = fil_agen & " WHERE agencia.tenro=28 "
            fil_agen = fil_agen & " AND (agencia.htetdesde <= " & ConvFecha(l_hasta) & " "
            fil_agen = fil_agen & " AND (agencia.htethasta IS NULL OR agencia.htethasta >= " & ConvFecha(l_hasta) & ")) )"
    Else
        If agencia = -2 Then
            fil_agen = " AND empleado.ternro in (SELECT ternro from his_estructura agencia "
            fil_agen = fil_agen & " WHERE agencia.tenro=28 "
            fil_agen = fil_agen & " AND (agencia.htetdesde <= " & ConvFecha(l_hasta) & " "
            fil_agen = fil_agen & " AND (agencia.htethasta IS NULL OR agencia.htethasta >= " & ConvFecha(l_hasta) & ")) )"
        Else
            If agencia <> 0 Then 'este caso se da cuando selecionamos una agencia determinada
                fil_agen = " AND empleado.ternro in (SELECT ternro from his_estructura agencia "
                fil_agen = fil_agen & " WHERE agencia.tenro=28 AND agencia.estrnro=" & agencia & ""
                fil_agen = fil_agen & " AND (agencia.htetdesde <= " & ConvFecha(l_hasta) & " "
                fil_agen = fil_agen & " AND (agencia.htethasta IS NULL OR agencia.htethasta >= " & ConvFecha(l_hasta) & ")) )"
            End If
        End If
    End If

    '----------------------------------------------------------------------------------------
    ' Calculo las bajas que tuvo el tipo de estructura
    '----------------------------------------------------------------------------------------
    
    If tenro3 <> 0 Then ' esto ocurre solo cuando se seleccionan los tres niveles
        StrSql = "SELECT COUNT (DISTINCT empleado.ternro) cantemplbaja "
        StrSql = StrSql & " FROM empleado "
        StrSql = StrSql & " INNER JOIN tercero ON tercero.ternro = empleado.ternro "
        StrSql = StrSql & " INNER JOIN fases ON fases.empleado = empleado.ternro"
        StrSql = StrSql & " LEFT JOIN causa ON fases.caunro = causa.caunro"
        StrSql = StrSql & " INNER JOIN his_estructura estact1 ON empleado.ternro = estact1.ternro  AND estact1.tenro  = " & tenro1 & ""
        StrSql = StrSql & " AND (estact1.htetdesde <=" & ConvFecha(l_hasta) & " AND (estact1.htethasta IS NULL OR estact1.htethasta >=" & ConvFecha(l_hasta) & "))"

        If estrnro1 <> 0 Then 'cuando se le asigna un valor al nivel 1
            StrSql = StrSql & " AND estact1.estrnro =" & estrnro1
        End If
        
        StrSql = StrSql & " INNER JOIN his_estructura estact2 ON empleado.ternro = estact2.ternro  AND estact2.tenro  = " & tenro2 & ""
        StrSql = StrSql & " AND (estact2.htetdesde<=" & ConvFecha(l_hasta) & " AND (estact2.htethasta IS NULL OR estact2.htethasta>=" & ConvFecha(l_hasta) & "))"
        
        If estrnro2 <> 0 Then 'cuando se le asigna un valor al nivel 2
            StrSql = StrSql & " AND estact2.estrnro =" & estrnro2
        End If
        
        StrSql = StrSql & " INNER JOIN his_estructura estact3 ON empleado.ternro = estact3.ternro  AND estact3.tenro  = " & tenro3 & ""
        StrSql = StrSql & " AND (estact3.htetdesde<=" & ConvFecha(l_hasta) & " AND (estact3.htethasta IS NULL OR estact3.htethasta>=" & ConvFecha(l_hasta) & "))"
        
        If estrnro3 <> 0 Then 'cuando se le asigna un valor al nivel 3
            StrSql = StrSql & " AND estact3.estrnro =" & estrnro3
        End If
        
        StrSql = StrSql & " WHERE fases.bajfec >= " & ConvFecha(l_desde) & " AND fases.bajfec <= " & ConvFecha(l_hasta)
        StrSql = StrSql & " " & fil_agen & ""
        
        
    Else
        If tenro2 <> 0 Then ' ocurre cuando se selecciono hasta el segundo nivel
            StrSql = "SELECT COUNT (DISTINCT empleado.ternro) cantemplbaja "
            StrSql = StrSql & " FROM empleado "
            StrSql = StrSql & " INNER JOIN tercero ON tercero.ternro = empleado.ternro "
            StrSql = StrSql & " INNER JOIN fases ON fases.empleado = empleado.ternro"
            StrSql = StrSql & " LEFT JOIN causa ON fases.caunro = causa.caunro"
            StrSql = StrSql & " INNER JOIN his_estructura estact1 ON empleado.ternro = estact1.ternro  AND estact1.tenro  = " & tenro1 & ""
            StrSql = StrSql & " AND (estact1.htetdesde <=" & ConvFecha(l_hasta) & " AND (estact1.htethasta IS NULL OR estact1.htethasta >=" & ConvFecha(l_hasta) & "))"

            If estrnro1 <> 0 Then
                StrSql = StrSql & " AND estact1.estrnro =" & estrnro1
            End If
            
            StrSql = StrSql & " INNER JOIN his_estructura estact2 ON empleado.ternro = estact2.ternro  AND estact2.tenro  = " & tenro2 & ""
            StrSql = StrSql & " AND (estact2.htetdesde<=" & ConvFecha(l_hasta) & " AND (estact2.htethasta IS NULL OR estact2.htethasta>=" & ConvFecha(l_hasta) & "))"
            
            If estrnro2 <> 0 Then
                StrSql = StrSql & " AND estact2.estrnro =" & estrnro2
            End If
            
            StrSql = StrSql & " WHERE fases.bajfec >= " & ConvFecha(l_desde) & " AND fases.bajfec <= " & ConvFecha(l_hasta)
            StrSql = StrSql & " " & fil_agen & ""
            
        Else
            If tenro1 <> 0 Then ' Cuando solo seleccionamos el primer nivel
                 StrSql = "SELECT COUNT (DISTINCT empleado.ternro) cantemplbaja "
                 StrSql = StrSql & " FROM empleado "
                 StrSql = StrSql & " INNER JOIN tercero ON tercero.ternro = empleado.ternro "
                 StrSql = StrSql & " INNER JOIN fases ON fases.empleado = empleado.ternro"
                 StrSql = StrSql & " LEFT JOIN causa ON fases.caunro = causa.caunro"
                 StrSql = StrSql & " INNER JOIN his_estructura estact1 ON empleado.ternro = estact1.ternro  AND estact1.tenro  = " & tenro1 & ""
                 StrSql = StrSql & " AND (estact1.htetdesde <=" & ConvFecha(l_hasta) & " AND (estact1.htethasta IS NULL OR estact1.htethasta >=" & ConvFecha(l_hasta) & "))"
   
                If estrnro1 <> 0 Then
                    StrSql = StrSql & " AND estact1.estrnro =" & estrnro1
                End If
                
                StrSql = StrSql & " WHERE fases.bajfec >= " & ConvFecha(l_desde) & " AND fases.bajfec <= " & ConvFecha(l_hasta)
                StrSql = StrSql & " " & fil_agen & ""
        
            Else ' cuando no hay nivel de estructura seleccionado
                    StrSql = "SELECT COUNT (DISTINCT empleado.ternro) cantemplbaja "
                    StrSql = StrSql & " FROM empleado "
                    StrSql = StrSql & " INNER JOIN tercero ON tercero.ternro = empleado.ternro "
                    StrSql = StrSql & " INNER JOIN fases ON fases.empleado = empleado.ternro"
                    StrSql = StrSql & " LEFT JOIN causa ON fases.caunro = causa.caunro"
                    StrSql = StrSql & " WHERE fases.bajfec >= " & ConvFecha(l_desde) & " AND fases.bajfec <= " & ConvFecha(l_hasta)
                    StrSql = StrSql & " " & fil_agen & ""
            
            End If
        End If
    End If
    
    'Flog.writeline " query bajas: " & StrSql
    
    'Flog.writeline "query Bajas Tiempo: " & Timer
        
    OpenRecordset StrSql, rs2
    If Not rs2.EOF Then
        EmplBajas = rs2!cantemplbaja
    Else
        EmplBajas = 0
    End If
    rs2.Close
    
Exit Function

ME_fases:
    Flog.writeline "    Error: " & Err.Description
    Flog.writeline "    SQL Ejecutado: " & StrSql
    Flog.writeline "  "
    
End Function

Function EmplActivos(l_desde, l_hasta)
Dim rs2 As New ADODB.Recordset
Dim fil_agen As String
Dim fil_empresa As String


On Error GoTo ME_fases

    fil_agen = "" ' cuando queremos todos los empleados
    If agencia = -1 Then
            fil_agen = " AND empleado.ternro not in (SELECT ternro from his_estructura agencia "
            fil_agen = fil_agen & " WHERE agencia.tenro=28 "
            fil_agen = fil_agen & " AND (agencia.htetdesde <= " & ConvFecha(l_hasta) & " "
            fil_agen = fil_agen & " AND (agencia.htethasta IS NULL OR agencia.htethasta >= " & ConvFecha(l_hasta) & ")) )"
    Else
        If agencia = -2 Then
            fil_agen = " AND empleado.ternro in (SELECT ternro from his_estructura agencia "
            fil_agen = fil_agen & " WHERE agencia.tenro=28 "
            fil_agen = fil_agen & " AND (agencia.htetdesde <= " & ConvFecha(l_hasta) & " "
            fil_agen = fil_agen & " AND (agencia.htethasta IS NULL OR agencia.htethasta >= " & ConvFecha(l_hasta) & ")) )"
        Else
            If agencia <> 0 Then 'este caso se da cuando selecionamos una agencia determinada
                fil_agen = " AND empleado.ternro in (SELECT ternro from his_estructura agencia "
                fil_agen = fil_agen & " WHERE agencia.tenro=28 AND agencia.estrnro=" & agencia & ""
                fil_agen = fil_agen & " AND (agencia.htetdesde <= " & ConvFecha(l_hasta) & " "
                fil_agen = fil_agen & " AND (agencia.htethasta IS NULL OR agencia.htethasta >= " & ConvFecha(l_hasta) & ")) )"
            End If
        End If
    End If
    
    '----------------------------------------------------------------------------------------
    ' Calculo los ingresos que tuvo el tipo de estructura
    '----------------------------------------------------------------------------------------
    
    If tenro3 <> 0 Then ' esto ocurre solo cuando se seleccionan los tres niveles
        StrSql = "SELECT COUNT (DISTINCT empleado.ternro) cantemplact "
        StrSql = StrSql & " FROM empleado "
        StrSql = StrSql & " INNER JOIN tercero ON tercero.ternro = empleado.ternro "
        StrSql = StrSql & " INNER JOIN his_estructura estact1 ON empleado.ternro = estact1.ternro  AND estact1.tenro  = " & tenro1 & ""
        StrSql = StrSql & " AND (estact1.htetdesde <=" & ConvFecha(l_hasta) & " AND (estact1.htethasta IS NULL OR estact1.htethasta >=" & ConvFecha(l_hasta) & "))"

        If estrnro1 <> 0 Then 'cuando se le asigna un valor al nivel 1
            StrSql = StrSql & " AND estact1.estrnro =" & estrnro1
        End If
        
        StrSql = StrSql & " INNER JOIN his_estructura estact2 ON empleado.ternro = estact2.ternro  AND estact2.tenro  = " & tenro2 & ""
        StrSql = StrSql & " AND (estact2.htetdesde<=" & ConvFecha(l_hasta) & " AND (estact2.htethasta IS NULL OR estact2.htethasta >=" & ConvFecha(l_hasta) & "))"
        
        If estrnro2 <> 0 Then 'cuando se le asigna un valor al nivel 2
            StrSql = StrSql & " AND estact2.estrnro =" & estrnro2
        End If
        
        StrSql = StrSql & " INNER JOIN his_estructura estact3 ON empleado.ternro = estact3.ternro  AND estact3.tenro  = " & tenro3 & ""
        StrSql = StrSql & " AND (estact3.htetdesde<=" & ConvFecha(l_hasta) & " AND (estact3.htethasta IS NULL OR estact3.htethasta >=" & ConvFecha(l_hasta) & "))"
        
        If estrnro3 <> 0 Then 'cuando se le asigna un valor al nivel 3
            StrSql = StrSql & " AND estact3.estrnro =" & estrnro3
        End If
        
        StrSql = StrSql & " WHERE (empleado.ternro in (SELECT empleado from fases WHERE fases.empleado = empleado.ternro AND (fases.altfec >=" & ConvFecha(l_desde) & " AND fases.altfec <=" & ConvFecha(l_hasta) & "))) "
        StrSql = StrSql & " " & fil_agen & ""
        
    Else
        If tenro2 <> 0 Then ' ocurre cuando se selecciono hasta el segundo nivel
            StrSql = "SELECT COUNT (DISTINCT empleado.ternro) cantemplact "
            StrSql = StrSql & " FROM empleado "
            StrSql = StrSql & " INNER JOIN tercero ON tercero.ternro = empleado.ternro "
            StrSql = StrSql & " INNER JOIN his_estructura estact1 ON empleado.ternro = estact1.ternro  AND estact1.tenro  = " & tenro1 & ""
            StrSql = StrSql & " AND (estact1.htetdesde <=" & ConvFecha(l_hasta) & " AND (estact1.htethasta IS NULL OR estact1.htethasta >=" & ConvFecha(l_hasta) & "))"

            If estrnro1 <> 0 Then
                StrSql = StrSql & " AND estact1.estrnro =" & estrnro1
            End If
            
            StrSql = StrSql & " INNER JOIN his_estructura estact2 ON empleado.ternro = estact2.ternro  AND estact2.tenro  = " & tenro2 & ""
            StrSql = StrSql & " AND (estact2.htetdesde<=" & ConvFecha(l_hasta) & " AND (estact2.htethasta IS NULL OR estact2.htethasta>=" & ConvFecha(l_hasta) & "))"
            
            If estrnro2 <> 0 Then
                StrSql = StrSql & " AND estact2.estrnro =" & estrnro2
            End If
            
            StrSql = StrSql & " WHERE (empleado.ternro in (SELECT empleado from fases WHERE fases.empleado = empleado.ternro AND (fases.altfec >=" & ConvFecha(l_desde) & " AND fases.altfec <=" & ConvFecha(l_hasta) & "))) "
            StrSql = StrSql & " " & fil_agen & ""
            
        Else
            If tenro1 <> 0 Then ' Cuando solo seleccionamos el primer nivel
                StrSql = "SELECT COUNT (DISTINCT empleado.ternro) cantemplact "
                StrSql = StrSql & " FROM empleado "
                StrSql = StrSql & " INNER JOIN tercero ON tercero.ternro = empleado.ternro "
                StrSql = StrSql & " INNER JOIN his_estructura estact1 ON empleado.ternro = estact1.ternro  AND estact1.tenro  = " & tenro1 & ""
                StrSql = StrSql & " AND (estact1.htetdesde <=" & ConvFecha(l_hasta) & " AND (estact1.htethasta IS NULL OR estact1.htethasta >=" & ConvFecha(l_hasta) & "))"
     
                If estrnro1 <> 0 Then
                    StrSql = StrSql & " AND estact1.estrnro =" & estrnro1
                End If
                
                StrSql = StrSql & " WHERE (empleado.ternro in (SELECT empleado from fases WHERE fases.empleado = empleado.ternro AND (fases.altfec >=" & ConvFecha(l_desde) & " AND fases.altfec <=" & ConvFecha(l_hasta) & "))) "
                StrSql = StrSql & " " & fil_agen & ""
        
            Else ' cuando no hay nivel de estructura seleccionado
                StrSql = "SELECT COUNT (DISTINCT empleado.ternro) cantemplact "
                StrSql = StrSql & " FROM empleado "
                StrSql = StrSql & " INNER JOIN tercero ON tercero.ternro = empleado.ternro "
                StrSql = StrSql & " WHERE (empleado.ternro in (SELECT empleado from fases WHERE fases.empleado = empleado.ternro AND (fases.altfec >=" & ConvFecha(l_desde) & " AND fases.altfec <=" & ConvFecha(l_hasta) & "))) "
                StrSql = StrSql & " " & fil_agen & ""
            
            End If
        End If
    End If
    
    'Flog.writeline " query Activos: " & StrSql
    
    'Flog.writeline "query Activos: " & Timer
    
    OpenRecordset StrSql, rs2
    If Not rs2.EOF Then
        EmplActivos = rs2!cantemplact
    Else
        EmplActivos = 0
    End If
    rs2.Close
    

Exit Function

ME_fases:
    Flog.writeline "    Error: " & Err.Description
    Flog.writeline "    SQL Ejecutado: " & StrSql
    Flog.writeline "  "
    
End Function



'--------------------------------------------------------------------
' Se encarga de generar los datos
'--------------------------------------------------------------------
Sub ReporteEvolPersonal()

Dim StrSql As String
Dim rsConsult As New ADODB.Recordset

Dim rsEstrnro As New ADODB.Recordset

Dim sqlAux As String
Dim rsCant As New ADODB.Recordset
Dim fil_agen As String
Dim anoini As Integer
Dim anofin  As Integer
Dim anoaux As Integer
Dim diffanno As Integer

Dim mesini As Integer
Dim mesfin  As Integer
Dim mesaux As Integer
Dim diffmeses As Integer
Dim l_desde As String
Dim l_hasta As String
Dim dias As Integer
Dim l_monto As Double
Dim dotacion As Double
Dim GrabaEmpleado As Boolean

Dim I, j, k As Integer



'Variables donde se guardan los datos del INSERT final

On Error GoTo MError


'*********************************************************************
'Ciclo por todos los empleados seleccionados del periodo
'*********************************************************************
GrabaEmpleado = False
fil_agen = ""

anoini = Year(fechadesde)
anofin = Year(fechahasta)
diffanno = DateDiff("yyyy", fechadesde, fechahasta)

mesini = Month(fechadesde)
mesfin = Month(fechahasta)

'diffmeses = DateDiff("m", fechadesde, fechahasta)

anoaux = anoini
'mesaux = mesini

mesaux = 1
diffmeses = 11

For I = 0 To diffanno
    For j = 0 To diffmeses
        l_desde = "01" & "/" & Format(mesaux, "00") & "/" & anoaux
        dias = CantidadDias(l_desde)
        l_hasta = dias & "/" & Format(mesaux, "00") & "/" & anoaux
                        
        'Agregada esta condicion para que coloque la dotacion en 0 para los meses desde 1 hasta mes incio
        ' cuando el mes inicio sea mayor al mesaux
        If (mesini > mesaux And anoaux = anoini) Then
            dotacion = 0
        Else
            dotacion = EmplActivos(l_desde, l_hasta) + EmplBajas(l_desde, l_hasta) + EmplExistentes(l_desde, l_hasta)
        End If
        'Fin
        
        'Flog.writeline " query 1: " & StrSql
        
        GrabaEmpleado = True
        Flog.writeline " Se encontraron datos para el mes: " & mesaux & " y año: " & anoaux & ""
        
        If dotacion = 0 Then
            l_monto = 0
        Else
            l_monto = dotacion
        End If
        
        'Inserto en la tabla rep_evol_per_det
        StrSql = " INSERT INTO rep_evol_per_det "
        StrSql = StrSql & " (bpronro, repdetanio, repdetmes, repdetdot"
        StrSql = StrSql & ")"
        StrSql = StrSql & " VALUES "
        StrSql = StrSql & "(" & NroProceso
        StrSql = StrSql & "," & anoaux
        StrSql = StrSql & "," & mesaux
        StrSql = StrSql & "," & numberForSQL(l_monto)
        StrSql = StrSql & ")"
        'Flog.writeline " query 2: " & StrSql
        '------------------------------------------------------------------
        'Guardo los datos en la BD
        '------------------------------------------------------------------
        objConn.Execute StrSql, , adExecuteNoRecords
        
        Flog.writeline " Se Grabo el detalle del Reporte de Evolucion del Personal"
        
        If mesaux = 12 Then
            'mesaux = mesini
            mesaux = 1
            j = diffmeses
        Else
            mesaux = mesaux + 1
            If (mesaux > mesfin And anoaux = anofin) Then
                For k = mesaux To 12
                    'Inserto en la tabla rep_evol_per_det
                    StrSql = " INSERT INTO rep_evol_per_det "
                    StrSql = StrSql & " (bpronro, repdetanio, repdetmes, repdetdot"
                    StrSql = StrSql & ")"
                    StrSql = StrSql & " VALUES "
                    StrSql = StrSql & "(" & NroProceso
                    StrSql = StrSql & "," & anoaux
                    StrSql = StrSql & "," & k
                    StrSql = StrSql & "," & numberForSQL(0)
                    StrSql = StrSql & ")"
                    'Flog.writeline " query 2.1: " & StrSql
                    '------------------------------------------------------------------
                    'Guardo los datos en la BD
                    '------------------------------------------------------------------
                    objConn.Execute StrSql, , adExecuteNoRecords
                    
                Next k
                j = diffmeses
                Flog.writeline " Se termina de armar el detalle del reporte"
            End If
        End If
    Next j
    anoaux = anoaux + 1
Next I

If GrabaEmpleado Then

    ' Se realiza el insert en la tabla cabecera
    StrSql = " INSERT INTO rep_evol_per "
    StrSql = StrSql & " (bpronro, repdescabr, repdescext, repaniodesde, repmesdesde, "
    StrSql = StrSql & " repaniohasta, repmeshasta, tenro1, estrnro1, tenro2, estrnro2, tenro3, estrnro3"
    StrSql = StrSql & ")"
    StrSql = StrSql & " VALUES "
    StrSql = StrSql & "(" & NroProceso
    StrSql = StrSql & ",'" & Mid(tituloReporte, 1, 200) & "'"
    StrSql = StrSql & ",'" & tituloReporte & "'"
    StrSql = StrSql & "," & Year(fechadesde)
    'StrSql = StrSql & "," & Month(fechadesde)
    StrSql = StrSql & "," & 1
    StrSql = StrSql & "," & Year(fechahasta)
    StrSql = StrSql & "," & Month(fechahasta)
    StrSql = StrSql & "," & tenro1
    StrSql = StrSql & "," & estrnro1
    StrSql = StrSql & "," & tenro2
    StrSql = StrSql & "," & estrnro2
    StrSql = StrSql & "," & tenro3
    StrSql = StrSql & "," & estrnro3
    StrSql = StrSql & ")"
    
    'Flog.writeline " query 3: " & StrSql
             
    '------------------------------------------------------------------
    'Guardo los datos en la BD
    '------------------------------------------------------------------
    objConn.Execute StrSql, , adExecuteNoRecords
    
    Flog.writeline " Se Grabo el Reporte Evolucion del Personal"
End If

Exit Sub

MError:
    Flog.writeline "Error en el Reporte Evolucion del Personal: " & NroProceso & " Error: " & Err.Description
    Flog.writeline "Ultimo Sql Ejecutado: " & StrSql
    HuboErrores = True
    EmpErrores = True
    Exit Sub
End Sub



'--------------------------------------------------------------------
' Se encarga de generar un ResultSet de los empleados a cambiar
' si el RS es vacio significa que hay que aplicarlo sobre todos
'--------------------------------------------------------------------
Sub CargarEmpleados(ByVal NroProc, ByRef rsEmpl As ADODB.Recordset, ByVal empresa As Long)

Dim StrEmpl As String

    If NroProc > 0 Then
        StrEmpl = "SELECT * FROM batch_empleado "
        StrEmpl = StrEmpl & " WHERE bpronro = " & NroProc
        StrEmpl = StrEmpl & " ORDER BY progreso,estado"
    End If
   
    OpenRecordset StrEmpl, rsEmpl
    
    cantRegistros = rsEmpl.RecordCount
    totalEmpleados = cantRegistros
    
End Sub

Function numberForSQL(Str)
   
  numberForSQL = Replace(Str, ",", ".")

End Function


Function strForSQL(Str)
   
  If IsNull(Str) Then
     strForSQL = "NULL"
  Else
     strForSQL = Str
  End If

End Function



Public Function Calcular_Edad(ByVal Fecha As Date, ByVal Hasta As Date) As Integer
'...........................................................................
' Archivo       : edad.i                              fecha ini. : 20/01/92
' Nombre progr. :
' tipo programa : FGZ
' Descripcion   :
'...........................................................................
Dim años  As Integer
Dim ALaFecha As Date

    ALaFecha = C_Date(Hasta)
    
    años = Year(ALaFecha) - Year(Fecha)
    If Month(ALaFecha) < Month(Fecha) Then
       años = años - 1
    Else
        If Month(ALaFecha) = Month(Fecha) Then
            If Day(ALaFecha) < Day(Fecha) Then
                años = años - 1
            End If
        End If
    End If
    Calcular_Edad = años
End Function


Sub buscarDatosEmpresa(Empnroestr)

Dim StrSql As String
Dim rsConsult As New ADODB.Recordset

    empresa = ""
    emprTer = 0
    emprCuit = ""
    emprDire = ""
    
    ' -------------------------------------------------------------------------
    'Busco los datos Basicos de la Empresa
    ' -------------------------------------------------------------------------
    Flog.writeline "Buscando datos de la empresa"
    
    StrSql = "SELECT * FROM empresa WHERE Estrnro = " & Empnroestr
    OpenRecordset StrSql, rsConsult
    
    If rsConsult.EOF Then
       Flog.writeline "Error: Buscando datos de la empresa: al obtener el empleado"
       HuboErrores = True
    Else
        empresa = rsConsult!empnom
        emprTer = rsConsult!ternro
        Empnro = rsConsult!Empnro
    End If
    
    rsConsult.Close
            
    'Consulta para obtener el RUT de la empresa
    StrSql = "SELECT nrodoc FROM tercero " & _
             " INNER JOIN ter_doc ON (tercero.ternro = ter_doc .ternro and ter_doc.tidnro = 1)" & _
             " Where tercero.ternro =" & emprTer
    
    Flog.writeline "Buscando datos del RUT de la empresa"
    
    OpenRecordset StrSql, rsConsult
    
    If rsConsult.EOF Then
        Flog.writeline "No se encontró el RUT de la Empresa"
        emprCuit = "  "
    Else
        emprCuit = rsConsult!nrodoc
    End If
    rsConsult.Close
End Sub

