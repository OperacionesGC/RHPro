Attribute VB_Name = "repIndRotBase"
'Global Const Version = "1.00" ' Carmen Quintero
'Global Const FechaModificacion = "29/08/2012"
'Global Const UltimaModificacion = "" 'Version Inicial

'Global Const Version = "1.01" ' Carmen Quintero
'Global Const FechaModificacion = "04/09/2012"
'Global Const UltimaModificacion = "" 'Carmen Quintero (16170) Se modificó la manera de actualizar el progreso del proceso.

'Global Const Version = "1.02" ' Carmen Quintero
'Global Const FechaModificacion = "19/09/2012"
'Global Const UltimaModificacion = "" 'Carmen Quintero (16170) Se modificó la performan del proceso.

'Global Const Version = "1.03" ' Carmen Quintero
'Global Const FechaModificacion = "13/11/2012"
'Global Const UltimaModificacion = "" 'Carmen Quintero (16170) Desarrollo de mejoras.

Global Const Version = "1.04" ' Carmen Quintero
Global Const FechaModificacion = "19/03/2013"
Global Const UltimaModificacion = "" 'Carmen Quintero (16170) Se modificó la consulta de los empleados que se encuentran de baja
                                     'dado un determinado mes.


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

Global bajasmes As Double
Global bajastotal As Double

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
    
    Nombre_Arch = PathFLog & "ReporteIndRotBase" & "-" & NroProceso & ".log"
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
       
       'Obtengo la empresa que se configuró
       empresa = CLng(ArrParametros(1))
       Flog.writeline "Se selecciono el parametro Empresa. " & ArrParametros(1)
    
       
       'Obtengo el tipo de estructura 1 si se configuró
       tenro1 = CLng(ArrParametros(2))
       Flog.writeline "Se selecciono el parametro Tipo de Estructura 1. " & ArrParametros(2)
       
       
       'Obtengo la estructura 1 si se configuró
       estrnro1 = CLng(ArrParametros(3))
       Flog.writeline "Se selecciono el parametro Estructura 1. " & ArrParametros(3)
       
       
       'Obtengo el tipo de estructura 2 si se configuró
       tenro2 = CLng(ArrParametros(4))
       Flog.writeline "Se selecciono el parametro Tipo de Estructura 2. " & ArrParametros(4)
       
       
       'Obtengo la estructura 2 si se configuró
       estrnro2 = CLng(ArrParametros(5))
       Flog.writeline "Se selecciono el parametro Estructura 2. " & ArrParametros(5)
       
       
       'Obtengo la fecha desde
       fechadesde = ArrParametros(6)
       Flog.writeline "Se selecciono el parametro fecha desde. " & ArrParametros(6)
       If Len(fechadesde) = 0 Then
            Flog.writeline "No Se selecciono el parametro fecha desde. "
            HuboErrores = True
       End If

       'Obtengo la fecha hasta
       fechahasta = ArrParametros(7)
       Flog.writeline "Se selecciono el parametro fecha hasta. " & ArrParametros(7)
       If Len(fechahasta) = 0 Then
            Flog.writeline "No Se selecciono el parametro fecha hasta. "
            HuboErrores = True
       End If
           
       tituloReporte = ArrParametros(8)
       Flog.writeline "Se selecciono el parametro titulo del reporte. " & ArrParametros(8)
       
              
       'EMPIEZA EL PROCESO
       Flog.writeline "Generando el reporte"
                    
                  
       'Obtengo los empleados sobre los que tengo que generar el reporte
       'CargarEmpleados(ByVal NroProc, ByRef rsEmpl As ADODB.Recordset, ByVal empresa As Long)
       CargarEmpleados NroProceso, rsEmpl, 0
       If Not rsEmpl.EOF Then
            Flog.writeline "Cantidad de empleados a procesar: " & cantRegistros
            CantEmpGrabados = cantRegistros 'Cantidad de empleados Guardados
       Else
            Flog.writeline "No hay empleados para el filtro seleccionado."
            HuboErrores = True
       End If
    
       'Actualizo Barch Proceso
       StrSql = "UPDATE batch_proceso SET bprcprogreso = 1 " & _
                ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'" & _
                ", bprcempleados ='" & CStr(rsEmpl.RecordCount) & "' WHERE bpronro = " & NroProceso
    
       objConn.Execute StrSql, , adExecuteNoRecords
 
       'Verifico que batch_empleado tenga registros
       If Not rsEmpl.EOF Then
            EmpErrores = False
            Flog.writeline ""
            Flog.writeline "Generando datos de los empleados "
            
            Call ReporteIndRotacionBase
                                            
            'Borro batch empleado
            '****************************************************************
            StrSql = "DELETE  FROM batch_empleado "
            StrSql = StrSql & " WHERE bpronro = " & NroProceso
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
    Flog.writeline "Cantidad de empleados guardados en el reporte: " & CantEmpGrabados
    Flog.Close

    Exit Sub
    
CE:
    HuboErrores = True
    Flog.writeline " Error: " & Err.Description & Now
    Flog.writeline " Ultimo SQL: " & StrSql
End Sub


Function BajasCausas(Desde, Hasta, caunro)
Dim rs2 As New ADODB.Recordset
Dim fil_agen As String
Dim fil_empresa As String

On Error GoTo ME_fases

    fil_agen = "" ' cuando queremos todos los empleados
    fil_empresa = ""

    fil_agen = " AND empleado.ternro not in (SELECT ternro from his_estructura agencia "
    fil_agen = fil_agen & " WHERE agencia.tenro=28 "
    fil_agen = fil_agen & " AND (agencia.htetdesde <= " & ConvFecha(Hasta) & " "
    fil_agen = fil_agen & " AND (agencia.htethasta IS NULL OR agencia.htethasta >= " & ConvFecha(Hasta) & ")) )"

    fil_empresa = " AND empleado.ternro in (SELECT ternro from his_estructura empresa "
    fil_empresa = fil_empresa & " WHERE empresa.tenro=10 and empresa.estrnro=" & empresa & ""
    fil_empresa = fil_empresa & " AND (empresa.htetdesde <=" & ConvFecha(Hasta) & ""
    fil_empresa = fil_empresa & " AND (empresa.htethasta IS NULL OR empresa.htethasta >=" & ConvFecha(Hasta) & ")) )"
        
'----------------------------------------------------------------------------------------
' Calculo las bajas que tuvo el tipo de estructura
'----------------------------------------------------------------------------------------
    
    StrSql = "SELECT COUNT (DISTINCT empleado.ternro) cantemplbaja "
    StrSql = StrSql & " FROM empleado "
    StrSql = StrSql & " INNER JOIN tercero ON tercero.ternro = empleado.ternro "
    StrSql = StrSql & " INNER JOIN fases ON fases.empleado = empleado.ternro "
    StrSql = StrSql & " INNER JOIN causa ON fases.caunro = causa.caunro "
    StrSql = StrSql & " INNER JOIN his_estructura estact1 ON empleado.ternro = estact1.ternro  "
    StrSql = StrSql & " INNER JOIN his_estructura estact2 ON empleado.ternro = estact2.ternro  "
    
    StrSql = StrSql & " WHERE fases.bajfec >= " & ConvFecha(Desde) & " AND fases.bajfec <= " & ConvFecha(Hasta)
    StrSql = StrSql & " AND causa.caunro = " & caunro & ""
    StrSql = StrSql & " AND estact1.tenro  = " & tenro1 & ""
    StrSql = StrSql & " AND (estact1.htetdesde <=" & ConvFecha(Hasta) & " AND (estact1.htethasta IS NULL OR estact1.htethasta >=" & ConvFecha(Hasta) & "))"
    StrSql = StrSql & " AND estact2.tenro  = " & tenro2 & ""
    StrSql = StrSql & " AND (estact2.htetdesde <=" & ConvFecha(Hasta) & " AND (estact2.htethasta IS NULL OR estact2.htethasta >=" & ConvFecha(Hasta) & "))"
    
    StrSql = StrSql & fil_agen & fil_empresa
    
    'Flog.writeline " query causas bajas: " & StrSql
    
    'Flog.writeline " query causas Bajas Tiempo: " & Timer
       
    OpenRecordset StrSql, rs2
    
    If Not rs2.EOF Then
        BajasCausas = rs2!cantemplbaja
    Else
        BajasCausas = 0
    End If
    rs2.Close
    
Exit Function

ME_fases:
    Flog.writeline "    Error: " & Err.Description
    Flog.writeline "    SQL Ejecutado: " & StrSql
    Flog.writeline "  "
    
End Function

Sub CargarCausas(Anio, mes, Desde, Hasta, dottotal)
Dim rs2 As New ADODB.Recordset
Dim l_porc As Double
Dim fil_agen As String
Dim fil_empresa As String


On Error GoTo ME_conf

    Flog.writeline " Buscando las causas de bajas para la fecha: " & Desde & " " & Hasta

    fil_agen = "" ' cuando queremos todos los empleados
    fil_empresa = ""

    fil_agen = " AND empleado.ternro not in (SELECT ternro from his_estructura agencia "
    fil_agen = fil_agen & " WHERE agencia.tenro=28 "
    fil_agen = fil_agen & " AND (agencia.htetdesde <= " & ConvFecha(Hasta) & " "
    fil_agen = fil_agen & " AND (agencia.htethasta IS NULL OR agencia.htethasta >= " & ConvFecha(Hasta) & ")) )"

    fil_empresa = " AND empleado.ternro in (SELECT ternro from his_estructura empresa "
    fil_empresa = fil_empresa & " WHERE empresa.tenro=10 and empresa.estrnro=" & empresa & ""
    fil_empresa = fil_empresa & " AND (empresa.htetdesde <=" & ConvFecha(Hasta) & ""
    fil_empresa = fil_empresa & " AND (empresa.htethasta IS NULL OR empresa.htethasta >=" & ConvFecha(Hasta) & ")) )"
        
'----------------------------------------------------------------------------------------
' Calculo las bajas que tuvo el tipo de estructura
'----------------------------------------------------------------------------------------
    
'    StrSql = "SELECT causa.caunro, COUNT (DISTINCT empleado.ternro)cantemplbaja"
'    StrSql = StrSql & " FROM empleado "
'    StrSql = StrSql & " INNER JOIN tercero ON tercero.ternro = empleado.ternro "
'    StrSql = StrSql & " INNER JOIN fases ON fases.empleado = empleado.ternro "
'    StrSql = StrSql & " INNER JOIN causa ON fases.caunro = causa.caunro "
'    StrSql = StrSql & " INNER JOIN his_estructura estact1 ON empleado.ternro = estact1.ternro  "
'    StrSql = StrSql & " INNER JOIN his_estructura estact2 ON empleado.ternro = estact2.ternro  "
'
'    StrSql = StrSql & " WHERE fases.bajfec >= " & ConvFecha(Desde) & " AND fases.bajfec <= " & ConvFecha(Hasta)
'    StrSql = StrSql & " AND estact1.tenro  = " & tenro1 & ""
'    StrSql = StrSql & " AND (estact1.htetdesde <=" & ConvFecha(Hasta) & " AND (estact1.htethasta IS NULL OR estact1.htethasta >=" & ConvFecha(Hasta) & "))"
'    StrSql = StrSql & " AND estact2.tenro  = " & tenro2 & ""
'    StrSql = StrSql & " AND (estact2.htetdesde <=" & ConvFecha(Hasta) & " AND (estact2.htethasta IS NULL OR estact2.htethasta >=" & ConvFecha(Hasta) & "))"
'
'    StrSql = StrSql & fil_agen & fil_empresa
'    StrSql = StrSql & " AND empleado.ternro IN (SELECT ternro FROM his_estructura causabaja "
'    StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = causabaja.estrnro "
'    StrSql = StrSql & " WHERE causabaja.tenro = 50 "
'    StrSql = StrSql & " AND (causabaja.htetdesde <= " & ConvFecha(Hasta) & " "
'    StrSql = StrSql & " AND (causabaja.htethasta IS NULL OR causabaja.htethasta >= " & ConvFecha(Hasta) & ")) "
'    StrSql = StrSql & " AND estructura.estrcodext in ('DES','REN') ) "
'    StrSql = StrSql & " GROUP BY causa.caunro "
'    StrSql = StrSql & " ORDER BY causa.caunro "

    StrSql = "SELECT causabaja.estrnro, COUNT (DISTINCT empleado.ternro)cantemplbaja"
    StrSql = StrSql & " FROM fases "
    StrSql = StrSql & " LEFT JOIN causa ON fases.caunro = causa.caunro "
    StrSql = StrSql & " INNER JOIN empleado ON fases.empleado = empleado.ternro "
    StrSql = StrSql & " INNER JOIN tercero ON tercero.ternro = empleado.ternro "
    StrSql = StrSql & " INNER JOIN his_estructura causabaja ON causabaja.ternro = empleado.ternro "
    StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = causabaja.estrnro "
    StrSql = StrSql & " WHERE fases.bajfec >= " & ConvFecha(Desde) & " AND fases.bajfec <= " & ConvFecha(Hasta)
    StrSql = StrSql & " AND estructura.estrcodext in ('DES','REN') "
    StrSql = StrSql & fil_agen & fil_empresa
    StrSql = StrSql & " GROUP BY causabaja.estrnro "
    StrSql = StrSql & " ORDER BY causabaja.estrnro "
    OpenRecordset StrSql, rs2
    
    While Not rs2.EOF
        'se calcula la rotacion sobre el total
        If dottotal > 0 Then
            l_porc = Round(((rs2("cantemplbaja") / dottotal) * 100), 2)
        Else
            l_porc = 0
        End If
        
        'Inserto en la tabla rep_ind_rot_sal_det
        StrSql = " INSERT INTO rep_ind_rot_sal_det "
        StrSql = StrSql & " (bpronro, repdetanio, repdetmes, caunro, repdetpor"
        StrSql = StrSql & ")"
        StrSql = StrSql & " VALUES "
        StrSql = StrSql & "(" & NroProceso
        StrSql = StrSql & "," & Anio
        StrSql = StrSql & "," & mes
        StrSql = StrSql & "," & rs2("estrnro")
        StrSql = StrSql & "," & numberForSQL(l_porc)
        StrSql = StrSql & ")"
        
        'Flog.writeline " query : " & StrSql
        '------------------------------------------------------------------
        'Guardo los datos en la BD
        '------------------------------------------------------------------
        objConn.Execute StrSql, , adExecuteNoRecords
        rs2.MoveNext
    Wend
    rs2.Close
Exit Sub

ME_conf:
   
    Flog.writeline "    Error: " & Err.Description
    Flog.writeline "    SQL Ejecutado: " & StrSql

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

Function EmplBajas(Desde, Hasta, genestrnro, secestrnro)
Dim rs2 As New ADODB.Recordset
Dim rs3 As New ADODB.Recordset
Dim fil_agen As String
Dim fil_empresa As String
Dim fil_fase As String
Dim l_valor As Double
Dim l_fechaanterior As String


On Error GoTo ME_fases

    fil_agen = "" ' cuando queremos todos los empleados
    fil_empresa = ""
    fil_fase = ""
    l_valor = 0

    
    fil_agen = " AND empleado.ternro not in (SELECT ternro from his_estructura agencia "
    fil_agen = fil_agen & " WHERE agencia.tenro=28 "
    fil_agen = fil_agen & " AND (agencia.htetdesde <= " & ConvFecha(Hasta) & " "
    fil_agen = fil_agen & " AND (agencia.htethasta IS NULL OR agencia.htethasta >= " & ConvFecha(Hasta) & ")) )"

    fil_empresa = " AND empleado.ternro in (SELECT ternro from his_estructura empresa "
    fil_empresa = fil_empresa & " WHERE empresa.tenro=10 and empresa.estrnro=" & empresa & ""
    fil_empresa = fil_empresa & " AND (empresa.htetdesde <=" & ConvFecha(Hasta) & ""
    fil_empresa = fil_empresa & " AND (empresa.htethasta IS NULL OR empresa.htethasta >=" & ConvFecha(Hasta) & ")) )"
    
    fil_fase = " AND empleado.ternro in (SELECT empleado FROM fases activa "
    fil_fase = fil_fase & " INNER JOIN tercero ON tercero.ternro = activa.empleado "
    fil_fase = fil_fase & " WHERE activa.empleado = empleado.ternro "
    
    If TipoBD = 3 Then ' Sqlserver
        fil_fase = fil_fase & " AND (activa.altfec <=  DateAdd(day, -1, fases.bajfec) "
        fil_fase = fil_fase & " AND (activa.bajfec IS NULL OR activa.bajfec >= DateAdd(day, -1, fases.bajfec) ))  "
        fil_fase = fil_fase & " AND MONTH("
        fil_fase = fil_fase & " DateAdd(day, -1, fases.bajfec) )= MONTH(" & ConvFecha(Hasta) & "))"
    End If
    
    If TipoBD = 4 Then ' Oracle
        fil_fase = fil_fase & " AND (activa.altfec <=  to_date(fases.bajfec,'dd/mm/yyyy')-1 "
        fil_fase = fil_fase & " AND (activa.bajfec IS NULL OR activa.bajfec >= to_date(fases.bajfec,'dd/mm/yyyy')-1 )) "
        fil_fase = fil_fase & " AND extract(month from to_date(fases.bajfec,'dd/mm/yyyy')-1)=extract(month from to_date(" & ConvFecha(Hasta) & ")))"
    End If
    
'----------------------------------------------------------------------------------------
' Calculo las bajas que tuvo el tipo de estructura
'----------------------------------------------------------------------------------------
    
    StrSql = "SELECT COUNT(DISTINCT empleado.ternro) cantemplbaja, fases.bajfec  "
    StrSql = StrSql & " FROM empleado "
    StrSql = StrSql & " INNER JOIN tercero ON tercero.ternro = empleado.ternro "
    StrSql = StrSql & " INNER JOIN fases ON fases.empleado = empleado.ternro"
    StrSql = StrSql & " INNER JOIN causa ON fases.caunro = causa.caunro"
    StrSql = StrSql & " INNER JOIN his_estructura estact1 ON empleado.ternro = estact1.ternro  AND estact1.tenro  = " & tenro1 & ""
      
    If secestrnro <> 0 Then
        StrSql = StrSql & " INNER JOIN his_estructura estact2 ON empleado.ternro = estact2.ternro  AND estact2.tenro  = " & tenro2 & ""
    End If

    StrSql = StrSql & " WHERE fases.bajfec >= " & ConvFecha(Desde) & " AND fases.bajfec <= " & ConvFecha(Hasta)
    StrSql = StrSql & " AND (estact1.htetdesde <=" & ConvFecha(Hasta) & " AND (estact1.htethasta IS NULL OR estact1.htethasta >=" & ConvFecha(Hasta) & "))"
    
    If genestrnro <> 0 Then
        StrSql = StrSql & " AND estact1.estrnro =" & genestrnro
    End If

    If secestrnro <> 0 Then
        StrSql = StrSql & " AND (estact2.htetdesde <=" & ConvFecha(Hasta) & " AND (estact2.htethasta IS NULL OR estact2.htethasta >=" & ConvFecha(Hasta) & "))"
        StrSql = StrSql & " AND estact2.estrnro =" & secestrnro
    End If
    
    
    StrSql = StrSql & fil_agen & fil_empresa & fil_fase

    StrSql = StrSql & "GROUP BY fases.bajfec"
    
    'Flog.writeline " query bajas: " & StrSql
    
    'Flog.writeline "query Bajas Tiempo: " & Timer
        
    OpenRecordset StrSql, rs2
     While Not rs2.EOF
            l_valor = l_valor + rs2!cantemplbaja
        rs2.MoveNext
     Wend
    
'    If Not rs2.EOF Then
'        EmplBajas = rs2!cantemplbaja
'    Else
'        EmplBajas = 0
'    End If

    rs2.Close
    EmplBajas = l_valor
    
Exit Function

ME_fases:
    Flog.writeline "    Error: " & Err.Description
    Flog.writeline "    SQL Ejecutado: " & StrSql
    Flog.writeline "  "
    
End Function

Function EmplActivos(Desde, Hasta, genestrnro, secestrnro)
Dim rs2 As New ADODB.Recordset
Dim fil_agen As String
Dim fil_empresa As String


On Error GoTo ME_fases

    fil_agen = "" ' cuando queremos todos los empleados
    fil_empresa = ""

    fil_agen = " AND empleado.ternro not in (SELECT ternro from his_estructura agencia "
    fil_agen = fil_agen & " WHERE agencia.tenro=28 "
    fil_agen = fil_agen & " AND (agencia.htetdesde <= " & ConvFecha(Hasta) & " "
    fil_agen = fil_agen & " AND (agencia.htethasta IS NULL OR agencia.htethasta >= " & ConvFecha(Hasta) & ")) )"

    fil_empresa = " AND empleado.ternro in (SELECT ternro from his_estructura empresa "
    fil_empresa = fil_empresa & " WHERE empresa.tenro=10 and empresa.estrnro=" & empresa & ""
    fil_empresa = fil_empresa & " AND (empresa.htetdesde<=" & ConvFecha(Hasta) & ""
    fil_empresa = fil_empresa & " AND (empresa.htethasta IS NULL OR empresa.htethasta>=" & ConvFecha(Hasta) & ")) )"

    '----------------------------------------------------------------------------------------
    ' Calculo los ingresos que tuvo el tipo de estructura
    '----------------------------------------------------------------------------------------
    
    StrSql = "SELECT COUNT (DISTINCT empleado.ternro) cantemplact "
    StrSql = StrSql & " FROM empleado "
    StrSql = StrSql & " INNER JOIN tercero ON tercero.ternro = empleado.ternro "
    StrSql = StrSql & " INNER JOIN his_estructura estact1 ON empleado.ternro = estact1.ternro  AND estact1.tenro  = " & tenro1 & ""

    If secestrnro <> 0 Then
        StrSql = StrSql & " INNER JOIN his_estructura estact2 ON empleado.ternro = estact2.ternro  AND estact2.tenro  = " & tenro2 & ""
    End If
    
    StrSql = StrSql & " WHERE (empleado.ternro in (SELECT empleado from fases WHERE fases.empleado = empleado.ternro AND (fases.altfec <=" & ConvFecha(Hasta) & " AND (fases.bajfec is null or fases.bajfec >=" & ConvFecha(Hasta) & ")))) "
    StrSql = StrSql & " AND (estact1.htetdesde <=" & ConvFecha(Hasta) & " AND (estact1.htethasta IS NULL OR estact1.htethasta >=" & ConvFecha(Hasta) & "))"
    
    If genestrnro <> 0 Then
        StrSql = StrSql & " AND estact1.estrnro =" & genestrnro
    End If
    
    If secestrnro <> 0 Then
        StrSql = StrSql & " AND (estact2.htetdesde <=" & ConvFecha(Hasta) & " AND (estact2.htethasta IS NULL OR estact2.htethasta>=" & ConvFecha(Hasta) & "))"
        StrSql = StrSql & " AND estact2.estrnro =" & secestrnro
    End If
    
    StrSql = StrSql & fil_agen & fil_empresa
    
    'Flog.writeline " query Activos: " & StrSql
    
    'Flog.writeline " query Activos Tiempo: " & Timer
    
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

Function EmplExistentes(Desde, Hasta, genestrnro, secestrnro)
Dim rs2 As New ADODB.Recordset
Dim fil_agen As String
Dim fil_empresa As String


On Error GoTo ME_Empleados

    fil_agen = "" ' cuando queremos todos los empleados
    fil_empresa = ""

    fil_agen = " AND empleado.ternro not in (SELECT ternro from his_estructura agencia "
    fil_agen = fil_agen & " WHERE agencia.tenro=28 "
    fil_agen = fil_agen & " AND (agencia.htetdesde <= " & ConvFecha(Hasta) & " "
    fil_agen = fil_agen & " AND (agencia.htethasta IS NULL OR agencia.htethasta >= " & ConvFecha(Hasta) & ")) )"

    fil_empresa = " AND empleado.ternro in (SELECT ternro from his_estructura empresa "
    fil_empresa = fil_empresa & " WHERE empresa.tenro=10 and empresa.estrnro=" & empresa & ""
    fil_empresa = fil_empresa & " AND (empresa.htetdesde<=" & ConvFecha(Hasta) & ""
    fil_empresa = fil_empresa & " AND (empresa.htethasta IS NULL OR empresa.htethasta>=" & ConvFecha(Hasta) & ")) )"

    '----------------------------------------------------------------------------------------
    ' Calculo los ingresos que tuvo el tipo de estructura
    '----------------------------------------------------------------------------------------
    
    StrSql = "SELECT COUNT (DISTINCT empleado.ternro) cantexisten "
    StrSql = StrSql & " FROM empleado "
    StrSql = StrSql & " INNER JOIN tercero ON tercero.ternro = empleado.ternro "
    StrSql = StrSql & " INNER JOIN his_estructura estact1 ON empleado.ternro = estact1.ternro  AND estact1.tenro  = " & tenro1 & ""

    If secestrnro <> 0 Then
        StrSql = StrSql & " INNER JOIN his_estructura estact2 ON empleado.ternro = estact2.ternro  AND estact2.tenro  = " & tenro2 & ""
    End If
    
    StrSql = StrSql & " WHERE (estact1.htetdesde <=" & ConvFecha(Hasta) & " AND (estact1.htethasta IS NULL OR estact1.htethasta >=" & ConvFecha(Hasta) & "))"
    
    If genestrnro <> 0 Then
        StrSql = StrSql & " AND estact1.estrnro =" & genestrnro
    End If
    
    If secestrnro <> 0 Then
        StrSql = StrSql & " AND (estact2.htetdesde <=" & ConvFecha(Hasta) & " AND (estact2.htethasta IS NULL OR estact2.htethasta >=" & ConvFecha(Hasta) & "))"
        StrSql = StrSql & " AND estact2.estrnro =" & secestrnro
    End If
    
    StrSql = StrSql & fil_agen & fil_empresa
    
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

Function dotacion(Desde, Hasta, genestrnro, secestrnro)
    Dim cantact As Double
    Dim cantbaja  As Double
    
    cantact = 0
    cantbaja = 0
    
    If (genestrnro > 0) Then
        cantact = EmplActivos(Desde, Hasta, genestrnro, secestrnro)
        cantbaja = EmplBajas(Desde, Hasta, genestrnro, secestrnro)
        bajasmes = cantbaja
    Else
        cantact = EmplActivos(Desde, Hasta, 0, 0)
        cantbaja = EmplBajas(Desde, Hasta, 0, 0)
        bajastotal = cantbaja
    End If
    
    dotacion = cantact + cantbaja
    
End Function

' ___________________________________________________________________________________________________
' procedimiento que inserta los dato cabecera en la tabla
' ___________________________________________________________________________________________________
Sub InsertarDatosCab()
Dim Campos As String
Dim Valores As String

On Error GoTo MError

Flog.writeline " "

Campos = " (bpronro, repdesabr, repdesext, empnro, repanio, repmesdesde, "
Campos = Campos & " repmeshasta, tenro1, estrnro1, tenro2, estrnro2 )"

Valores = "("
Valores = Valores & NroProceso & ",'" & Mid(tituloReporte, 1, 200) & "','" & tituloReporte & "', " & empresa & " , "
Valores = Valores & Year(fechadesde) & "," & Month(fechadesde) & ", " & Month(fechahasta) & " , "
Valores = Valores & tenro1 & "," & estrnro1 & "," & tenro2 & "," & estrnro2 & ""
Valores = Valores & ")"

StrSql = " INSERT INTO rep_ind_rot " & Campos & " VALUES " & Valores

'Flog.writeline " query 4: " & StrSql

objConn.Execute StrSql, , adExecuteNoRecords

'Flog.writeline " Se Grabo la cabecera del Reporte Indicador Rotacion Base"

Exit Sub


MError:
    Flog.writeline "    Error: " & Err.Description
    Flog.writeline "    SQL Ejecutado: " & StrSql

End Sub
' ____________________________________________________________
' procedimiento que inserta los datos en la tabla detalle
' ____________________________________________________________
Sub InsertarDatosDet(orden, estrnro1, tenro, Anio, mes, dotacion, salidas, rottotal, rotsector, estrnro2)
Dim Campos As String
Dim Valores As String

On Error GoTo MError


Campos = " (bpronro, repdetorden, tenro, estrnro1,  repdetanio,  repdetmes, "
Campos = Campos & " repdetdot, repdetsal , repdetrottot, repdetrotsec, estrnro2 "
Campos = Campos & " )"

Valores = "("
Valores = Valores & NroProceso & "," & orden & "," & tenro & "," & estrnro1 & "," & Anio & ","
Valores = Valores & mes & "," & dotacion & "," & salidas & "," & rottotal & "," & rotsector & "," & estrnro2 & ""
Valores = Valores & ")"

StrSql = " INSERT INTO rep_ind_rot_det " & Campos & " VALUES " & Valores

'Flog.writeline " query 0: " & StrSql

objConn.Execute StrSql, , adExecuteNoRecords

'Flog.writeline " Se Grabo el detalle del Reporte de Indicador Rotacion Base"

Exit Sub

MError:
    Flog.writeline "    Error: " & Err.Description
    Flog.writeline "    SQL Ejecutado: " & StrSql

End Sub

'--------------------------------------------------------------------
' Se encarga de generar los datos
'--------------------------------------------------------------------
Sub ReporteIndRotacionBase()

Dim StrSql As String
Dim rsConsult As New ADODB.Recordset
Dim rsConsult2 As New ADODB.Recordset

Dim sqlAux As String

Dim Anio As Integer
Dim anioant As Integer

Dim mesini As Integer
Dim mesfin  As Integer
Dim mesaux As Integer
Dim diffmeses As Integer
Dim l_desde As String
Dim l_hasta As String
Dim dias As Integer
Dim l_monto As Double
Dim l_dotmes As Double
Dim l_rotsec As Double
Dim GrabaEmpleado As Boolean
Dim l_orden As Integer
Dim l_ordtipoestr As Integer
Dim l_dottotal As Double
Dim l_rottotal As Double
Dim valor As Double
Dim aux As Double
Dim auxtotal As Double
Dim total As Double
Dim fil_agen As String
Dim fil_empresa As String
Dim fil_fase As String
Dim l_estraux As Long
Dim l_dotaciontotal As Double
Dim l_estrpadre As Long
Dim l_salidastotal As Double
Dim l_rottotalpadre As Double
Dim l_rotsecpadre As Double
Dim l_encontrado As Boolean
Dim l_paso As Boolean
Dim l_salidas As Double

Dim j As Integer


'Variables donde se guardan los datos del INSERT final

On Error GoTo MError


'*********************************************************************
'Ciclo por todos los empleados seleccionados del periodo
'*********************************************************************

Anio = Year(fechadesde)

mesini = Month(fechadesde)
mesfin = Month(fechahasta)

diffmeses = DateDiff("m", fechadesde, fechahasta)

mesaux = mesini

valor = (99 / (diffmeses + 1))
auxtotal = 0

For j = 0 To diffmeses
    aux = 0
    l_desde = "01" & "/" & Format(mesaux, "00") & "/" & Anio
    dias = CantidadDias(l_desde)
    l_hasta = dias & "/" & Format(mesaux, "00") & "/" & Anio
    
    fil_agen = "" ' cuando queremos todos los empleados
    fil_empresa = ""

    fil_agen = " AND empleado.ternro not in (SELECT ternro from his_estructura agencia "
    fil_agen = fil_agen & " WHERE agencia.tenro=28 "
    fil_agen = fil_agen & " AND (agencia.htetdesde <= " & ConvFecha(l_hasta) & " "
    fil_agen = fil_agen & " AND (agencia.htethasta IS NULL OR agencia.htethasta >= " & ConvFecha(l_hasta) & ")) )"

    fil_empresa = " AND empleado.ternro in (SELECT ternro from his_estructura empresa "
    fil_empresa = fil_empresa & " WHERE empresa.tenro=10 and empresa.estrnro=" & empresa & ""
    fil_empresa = fil_empresa & " AND (empresa.htetdesde<=" & ConvFecha(l_hasta) & ""
    fil_empresa = fil_empresa & " AND (empresa.htethasta IS NULL OR empresa.htethasta>=" & ConvFecha(l_hasta) & ")) )"
    
'    fil_fase = " AND empleado.ternro in (SELECT empleado FROM fases activa "
'    fil_fase = fil_fase & " INNER JOIN tercero ON tercero.ternro = activa.empleado "
'    fil_fase = fil_fase & " WHERE activa.empleado = empleado.ternro "
'    fil_fase = fil_fase & " AND (activa.altfec <=  DateAdd(day, -1, fases.bajfec) "
'    fil_fase = fil_fase & " AND (activa.bajfec IS NULL OR activa.bajfec >= DateAdd(day, -1, fases.bajfec) ))  "
'    fil_fase = fil_fase & " AND MONTH("
'    fil_fase = fil_fase & " DateAdd(day, -1, fases.bajfec) )= MONTH(" & ConvFecha(l_hasta) & "))"
    
    fil_fase = " AND empleado.ternro in (SELECT empleado FROM fases activa "
    fil_fase = fil_fase & " INNER JOIN tercero ON tercero.ternro = activa.empleado "
    fil_fase = fil_fase & " WHERE activa.empleado = empleado.ternro "
    
    If TipoBD = 3 Then ' Sqlserver
        fil_fase = fil_fase & " AND (activa.altfec <=  DateAdd(day, -1, fases.bajfec) "
        fil_fase = fil_fase & " AND (activa.bajfec IS NULL OR activa.bajfec >= DateAdd(day, -1, fases.bajfec) ))  "
        fil_fase = fil_fase & " AND MONTH("
        fil_fase = fil_fase & " DateAdd(day, -1, fases.bajfec) )= MONTH(" & ConvFecha(l_hasta) & "))"
    End If
    
    If TipoBD = 4 Then ' Oracle
        fil_fase = fil_fase & " AND (activa.altfec <=  to_date(fases.bajfec,'dd/mm/yyyy')-1 "
        fil_fase = fil_fase & " AND (activa.bajfec IS NULL OR activa.bajfec >= to_date(fases.bajfec,'dd/mm/yyyy')-1 )) "
        fil_fase = fil_fase & " AND extract(month from to_date(fases.bajfec,'dd/mm/yyyy')-1)=extract(month from to_date(" & ConvFecha(l_hasta) & ")))"
    End If
    
    Flog.writeline " Buscando los datos para la fecha: " & l_desde & " " & l_hasta
    
    'se inicializan las variables
    l_dottotal = 0
    bajastotal = 0
    l_rottotal = 0
    l_salidastotal = 0
    l_dotaciontotal = 0
    l_encontrado = False
    l_orden = 1
    
    l_ordtipoestr = 2
    
    'se calcula la dotacion total del mes
    l_dottotal = dotacion(l_desde, l_hasta, 0, 0)
        
    'se calcula la rotacion sobre el total
    If l_dottotal > 0 Then
       l_rottotal = Round(((bajastotal / l_dottotal) * 100), 2)
    Else
        l_rottotal = 0
    End If
    
    'Inserto en la tabla rep_ind_rot_det
    Call InsertarDatosDet(0, 0, tenro1, Anio, mesaux, numberForSQL(l_dottotal), numberForSQL(bajastotal), numberForSQL(l_rottotal), 0, 0)
    
    StrSql = "SELECT estact1.estrnro gerencia, estact2.estrnro sector, COUNT(DISTINCT empleado.ternro)cantemplact "
    StrSql = StrSql & " FROM empleado "
    StrSql = StrSql & " INNER JOIN tercero ON tercero.ternro = empleado.ternro "
    StrSql = StrSql & " INNER JOIN his_estructura estact1 ON empleado.ternro = estact1.ternro  AND estact1.tenro  = " & tenro1 & ""
    StrSql = StrSql & " INNER JOIN his_estructura estact2 ON empleado.ternro = estact2.ternro  AND estact2.tenro  = " & tenro2 & ""
    StrSql = StrSql & " WHERE (empleado.ternro in (SELECT empleado from fases WHERE fases.empleado = empleado.ternro AND (fases.altfec <=" & ConvFecha(l_hasta) & " AND (fases.bajfec is null or fases.bajfec >=" & ConvFecha(l_hasta) & ")))) "
    StrSql = StrSql & " AND (estact1.htetdesde <=" & ConvFecha(l_hasta) & " AND (estact1.htethasta IS NULL OR estact1.htethasta >=" & ConvFecha(l_hasta) & "))"
    StrSql = StrSql & " AND (estact2.htetdesde <=" & ConvFecha(l_hasta) & " AND (estact2.htethasta IS NULL OR estact2.htethasta >=" & ConvFecha(l_hasta) & "))"
    StrSql = StrSql & fil_agen & fil_empresa
    StrSql = StrSql & " GROUP BY estact1.estrnro, estact2.estrnro "
    StrSql = StrSql & " ORDER BY estact1.estrnro, estact2.estrnro "
    OpenRecordset StrSql, rsConsult
    
    StrSql = "SELECT estact1.estrnro gerencia,estact2.estrnro sector,COUNT(DISTINCT empleado.ternro)cantemplbaja "
    StrSql = StrSql & " FROM empleado "
    StrSql = StrSql & " INNER JOIN tercero ON tercero.ternro = empleado.ternro "
    StrSql = StrSql & " INNER JOIN fases ON fases.empleado = empleado.ternro"
    StrSql = StrSql & " INNER JOIN causa ON fases.caunro = causa.caunro"
    StrSql = StrSql & " INNER JOIN his_estructura estact1 ON empleado.ternro = estact1.ternro  AND estact1.tenro  = " & tenro1 & ""
    StrSql = StrSql & " INNER JOIN his_estructura estact2 ON empleado.ternro = estact2.ternro  AND estact2.tenro  = " & tenro2 & ""
    StrSql = StrSql & " WHERE fases.bajfec >= " & ConvFecha(l_desde) & " AND fases.bajfec <= " & ConvFecha(l_hasta)
    StrSql = StrSql & " AND (estact1.htetdesde <=" & ConvFecha(l_hasta) & " AND (estact1.htethasta IS NULL OR estact1.htethasta >=" & ConvFecha(l_hasta) & "))"
    StrSql = StrSql & " AND (estact2.htetdesde <=" & ConvFecha(l_hasta) & " AND (estact2.htethasta IS NULL OR estact2.htethasta >=" & ConvFecha(l_hasta) & "))"
    StrSql = StrSql & fil_agen & fil_empresa & fil_fase
    StrSql = StrSql & " GROUP BY estact1.estrnro, estact2.estrnro"
    StrSql = StrSql & " ORDER BY estact1.estrnro, estact2.estrnro "
    OpenRecordset StrSql, rsConsult2
    
    l_estraux = 0
    
    While Not rsConsult.EOF
        If Not rsConsult2.EOF Then
            If rsConsult("gerencia") = rsConsult2("gerencia") And rsConsult("sector") = rsConsult2("sector") Then
                l_encontrado = True
            
                'se calcula la dotacion del mes por sector
                l_dotmes = rsConsult("cantemplact") + rsConsult2("cantemplbaja")
                l_salidas = l_salidas + rsConsult2("cantemplbaja")
                
                'se calcula la rotacion sobre el total
                If l_dottotal > 0 Then
                   l_rottotal = Round(((rsConsult2("cantemplbaja") / l_dottotal) * 100), 2)
                Else
                   l_rottotal = 0
                End If
    
                'se calcula la rotacion sobre el sector
                If l_dotmes > 0 Then
                   l_rotsec = Round(((rsConsult2("cantemplbaja") / l_dotmes) * 100), 2)
                Else
                   l_rotsec = 0
                End If
            Else
                l_dotmes = rsConsult("cantemplact")
                l_rottotal = 0
                l_rotsec = 0
                l_salidas = 0
            End If
        Else
            l_dotmes = rsConsult("cantemplact")
            l_rottotal = 0
            l_rotsec = 0
            l_salidas = 0
        End If
        
        Call InsertarDatosDet(l_ordtipoestr, rsConsult("gerencia"), tenro2, Anio, mesaux, numberForSQL(l_dotmes), numberForSQL(l_salidas), numberForSQL(l_rottotal), numberForSQL(l_rotsec), rsConsult("sector"))
        
        If l_estraux = 0 Then
           l_estraux = rsConsult("gerencia")
        End If
        
        If l_estraux = rsConsult("gerencia") Then
           'l_salidastotal = l_salidas
           l_salidastotal = l_salidastotal + l_salidas
           'l_dotaciontotal = l_dotaciontotal + rsConsult("cantemplact")
           l_dotaciontotal = l_dotaciontotal + l_dotmes
           l_estrpadre = rsConsult("gerencia")
                      
           If l_dottotal > 0 Then
              l_rottotalpadre = Round(((l_salidastotal / l_dottotal) * 100), 2)
           Else
              l_rottotalpadre = 0
           End If
           
           If l_dotaciontotal > 0 Then
              l_rotsecpadre = Round(((l_salidastotal / l_dotaciontotal) * 100), 2)
           Else
              l_rotsecpadre = 0
           End If
        Else
            'Inserto en la tabla rep_ind_rot_det
            Call InsertarDatosDet(l_orden, l_estrpadre, tenro1, Anio, mesaux, numberForSQL(l_dotaciontotal), numberForSQL(l_salidastotal), numberForSQL(l_rottotalpadre), numberForSQL(l_rotsecpadre), 0)
            l_orden = l_ordtipoestr - 1
            l_estraux = rsConsult("gerencia")
            l_dotaciontotal = l_dotmes
            l_estrpadre = l_estraux
            l_rottotalpadre = l_rottotal
            l_rotsecpadre = l_rotsec
            l_salidastotal = l_salidas
        End If
      
        rsConsult.MoveNext
        
        If Not rsConsult.EOF Then
            If l_estraux = rsConsult("gerencia") Then
               l_ordtipoestr = l_ordtipoestr + 1
            Else
               l_ordtipoestr = l_ordtipoestr + 2
            End If
        End If
    
        If l_encontrado Then
           rsConsult2.MoveNext
           l_encontrado = False
           l_salidas = 0
        End If
        
        'Actualizo el estado del proceso
        TiempoAcumulado = GetTickCount
    
        'Resto uno a la cantidad de registros
        cantRegistros = rsConsult.RecordCount
        
        total = (valor / cantRegistros)
        
        aux = aux + total + auxtotal
        
        auxtotal = 0
        
        'Actualizo
        StrSql = "UPDATE batch_proceso SET bprcprogreso = " & numberForSQL(aux) & _
               ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'" & _
               ", bprcempleados ='" & CStr(cantRegistros) & "' WHERE bpronro = " & NroProceso
        objConn.Execute StrSql, , adExecuteNoRecords
    Wend
    
    If rsConsult.EOF Then
        Call InsertarDatosDet(l_orden, l_estrpadre, tenro1, Anio, mesaux, numberForSQL(l_dotaciontotal), numberForSQL(l_salidastotal), numberForSQL(l_rottotalpadre), numberForSQL(l_rotsecpadre), 0)
    End If
    
    Call CargarCausas(Anio, mesaux, l_desde, l_hasta, l_dottotal)
    mesaux = mesaux + 1
    
    auxtotal = aux
Next j

' Se realiza el insert en la tabla cabecera
Call InsertarDatosCab
    
' Se calculan las causas de bajas del año anterior
anioant = Year(fechadesde) - 1
mesaux = mesini

For j = 0 To diffmeses
    l_desde = "01" & "/" & Format(mesaux, "00") & "/" & anioant
    dias = CantidadDias(l_desde)
    l_hasta = dias & "/" & Format(mesaux, "00") & "/" & anioant
    
    Flog.writeline " Buscando los datos para la fecha: " & l_desde & " " & l_hasta
    
    'se calcula la dotacion total del mes para el año anterior
    l_dottotal = dotacion(l_desde, l_hasta, 0, 0)
    
    Call CargarCausas(anioant, mesaux, l_desde, l_hasta, l_dottotal)
    
    mesaux = mesaux + 1
Next j

'Redondeo a 100%
If Int(aux) < 100 Then
    'Inserto progreso
    StrSql = "UPDATE batch_proceso SET bprcprogreso = 100"
    StrSql = StrSql & " WHERE bpronro = " & NroProceso
    objConn.Execute StrSql, , adExecuteNoRecords
End If

Flog.writeline ""
Flog.writeline "El proceso se realizó con éxito"

Exit Sub

MError:
    Flog.writeline "Error en el Reporte Indicador Rotacion Base: " & NroProceso & " Error: " & Err.Description
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

