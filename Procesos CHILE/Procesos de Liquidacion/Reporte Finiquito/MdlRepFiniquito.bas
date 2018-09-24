Attribute VB_Name = "MdlRepFiniquito"
Option Explicit

'Const Version = "1.0"
'Const FechaVersion = "03/11/2008"
'Autor = Martin Ferraro
'Const Version = "1.1"
'Const FechaVersion = "18/11/2008"
'Autor = Martin Ferraro
'Const Version = "1.2"
'Const FechaVersion = "31/07/2009"
'Autor = Martin Ferraro - Encriptacion de string connection

'Const Version = "1.3"
'Const FechaVersion = "12/11/2015"
'Autor = Dimatz Rafael - CAS 32780 - Se corrigio el Representante Legal, para que muestre el configurado por sistema

Const Version = "1.4"
Const FechaVersion = "16/11/2015"
'Autor = Dimatz Rafael - CAS 32780 - Se corrigio la query para que traiga bien los valores del Confrep

Global CantEmplError
Global CantEmplSinError
Global Errores As Boolean


Public Sub Main()
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento inicial
' Autor      : Martin Ferraro
' Fecha      : 03/11/2008
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------

Dim objconnMain As New ADODB.Connection
Dim strCmdLine
Dim Nombre_Arch As String
Dim rs_batch_proceso As New ADODB.Recordset
Dim PID As String
Dim bprcparam As String
Dim ArrParametros
       
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
    Nombre_Arch = PathFLog & "Rep_Finiquito" & "-" & NroProcesoBatch & ".log"
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
    If Err.Number <> 0 Or Error_Encrypt Then
        Flog.writeline "Problemas en la conexion"
        Exit Sub
    End If
    OpenConnection strconexion, objconnProgreso
    If Err.Number <> 0 Or Error_Encrypt Then
        Flog.writeline "Problemas en la conexion"
        Exit Sub
    End If
    On Error GoTo 0
    
    On Error GoTo ME_Main
    
    'Cambio el estado del proceso a Procesando
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcprogreso = 0 ,bprcestado = 'Procesando', bprcpid = " & PID & " WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    Flog.writeline "Pone el estado en procesando"
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE btprcnro = 232 AND bpronro =" & NroProcesoBatch
    OpenRecordset StrSql, rs_batch_proceso
    
    TiempoInicialProceso = GetTickCount
    
    If Not rs_batch_proceso.EOF Then
        bprcparam = rs_batch_proceso!bprcparam
        rs_batch_proceso.Close
        Set rs_batch_proceso = Nothing
        Call Finiquito(NroProcesoBatch, bprcparam)
    Else
        Flog.writeline "no encontró el proceso"
    End If
    
    TiempoFinalProceso = GetTickCount
    Flog.writeline
    Flog.writeline "**********************************************************"
    Flog.writeline
    Flog.writeline
    Flog.writeline "Tiempo del proceso (milisegundos): " & (TiempoFinalProceso - TiempoInicialProceso)
    Flog.writeline
    Flog.writeline "**********************************************************"
    If Not Errores Then
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado', bprcprogreso = 100 WHERE bpronro = " & NroProcesoBatch
    Else
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Error', bprcprogreso = 100  WHERE bpronro = " & NroProcesoBatch
    End If
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
Fin:
    Flog.Close
    'If objConn.State = adStateOpen Then objConn.Close
    If objconnProgreso.State = adStateOpen Then objconnProgreso.Close
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
    'MyBeginTrans
        StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcprogreso = 0, bprcestado = 'Error General', bprcpid = " & PID & " WHERE bpronro = " & NroProcesoBatch
        objconnProgreso.Execute StrSql, , adExecuteNoRecords
    'MyCommitTrans
End Sub


Public Sub Finiquito(ByVal bpronro As Long, ByVal Parametros As String)
' --------------------------------------------------------------------------------------------
' Descripcion: Procedimiento de generacion del reporte Finiquito
' Autor      : Martin Ferraro
' Fecha      : 04/11/2008
' --------------------------------------------------------------------------------------------

'-------------------------------------------------------------------------------------------------
'Recordsets
'-------------------------------------------------------------------------------------------------
Dim rs_Empleados As New ADODB.Recordset
Dim rs_Consult As New ADODB.Recordset

'-------------------------------------------------------------------------------------------------
'Parametros
'-------------------------------------------------------------------------------------------------
Dim legDesde As Long
Dim legHasta As Long
Dim Estado As Integer
Dim Empresa As Long
Dim pliqNro As Long
Dim proNro As Long


'-------------------------------------------------------------------------------------------------
'Variables
'-------------------------------------------------------------------------------------------------
Dim arrPar
Dim arrConc(11) As String
Dim arrAcu(11) As String
Dim arrEtiq(11) As String
Dim Ind As Long
Dim repLegal As String
Dim fechaPago As Date
Dim apeNom As String
Dim RUTEmpl As String
Dim domiEmpl As String
Dim faseDesde As String
Dim faseHasta As String
Dim causaBaja As String
Dim puesto As String
Dim empresaNom As String
Dim empTernro As Long
Dim RUTEmpr As String
Dim empresaProv As String
Dim Orden As Long
Dim Titulo As String
Dim totalConc As Double
Dim totalAcu As Double
Dim ordenFila As Long

On Error GoTo E_Finiquito

'-------------------------------------------------------------------------------------------------
' Levanto cada parametro por separado, el separador de parametros es "@"
'-------------------------------------------------------------------------------------------------
'l_legdesde & "@" & l_leghasta & "@" & l_estado & "@" & l_empresa & "@" & l_pliqnro & "@" & l_pronro
Flog.writeline "Levantando Parametros  "
Flog.writeline Espacios(Tabulador * 1) & Parametros
Flog.writeline

If Not IsNull(Parametros) Then
    
    If Len(Parametros) >= 1 Then
        
        arrPar = Split(Parametros, "@")
        If UBound(arrPar) = 5 Then
        
         'Legajo Desde
         '-----------------------------------------------------------
         legDesde = CLng(arrPar(0))
         Flog.writeline Espacios(Tabulador * 1) & "Parametro Legajo Desde = " & legDesde
         
         'Legajo Hasta
         '-----------------------------------------------------------
         legHasta = CLng(arrPar(1))
         Flog.writeline Espacios(Tabulador * 1) & "Parametro Legajo Hasta = " & legHasta
         
         'Estado
         '-----------------------------------------------------------
         Estado = CInt(arrPar(2))
         Flog.writeline Espacios(Tabulador * 1) & "Parametro Estado = " & Estado
         
         'Estado
         '-----------------------------------------------------------
         Empresa = CInt(arrPar(3))
         Flog.writeline Espacios(Tabulador * 1) & "Parametro Empresa = " & Empresa
         
         'Periodo
         '-----------------------------------------------------------
         pliqNro = CLng(arrPar(4))
         Flog.writeline Espacios(Tabulador * 1) & "Parametro Periodo = " & pliqNro
         
         'Proceso
         '-----------------------------------------------------------
         proNro = CLng(arrPar(5))
         Flog.writeline Espacios(Tabulador * 1) & "Parametro Proceso = " & proNro
         
        Else
            Flog.writeline Espacios(Tabulador * 0) & "ERROR. Numero de parametros erroneo."
            Exit Sub
        End If

    End If
Else
    Flog.writeline "ERROR. No se encontraron parametros para el proceso"
    Exit Sub
End If
Flog.writeline



'-------------------------------------------------------------------------------------------------
'Inicializacion de Variables
'-------------------------------------------------------------------------------------------------
For Ind = 0 To 10
    arrConc(Ind) = ""
    arrAcu(Ind) = ""
    arrEtiq(Ind) = ""
Next
repLegal = ""


'-------------------------------------------------------------------------------------------------
'Configuracion del Reporte
'-------------------------------------------------------------------------------------------------
Flog.writeline "Levantando configuracion del Reporte 247"
StrSql = "SELECT * FROM confrepadv WHERE repnro = 247 ORDER BY confnrocol"
OpenRecordset StrSql, rs_Consult
If rs_Consult.EOF Then
    Flog.writeline Espacios(Tabulador * 1) & "ERROR. No se encontró la configuración del Reporte"
    'Exit Sub
End If

Do While Not rs_Consult.EOF
    
    Select Case rs_Consult!confnrocol
        'Case 1:
            'Nombre del representante legal
            'repLegal = IIf(EsNulo(rs_Consult!confetiq), "", rs_Consult!confetiq)
        Case 2 To 12:
            'Conceptos y/o acumuladores a buscar
            Select Case UCase(rs_Consult!conftipo)
                Case "CO"
                    'Lista de Conceptos
                    If Len(arrConc(rs_Consult!confnrocol - 2)) = 0 Then
                        arrConc(rs_Consult!confnrocol - 2) = IIf(EsNulo(rs_Consult!confval), "", rs_Consult!confval)
                    Else
                        If Not EsNulo(rs_Consult!confval) Then
                            arrConc(rs_Consult!confnrocol - 2) = arrConc(rs_Consult!confnrocol - 2) & "," & rs_Consult!confval
                        End If
                    End If
                    
                    'Etiqueta de la columna
                    If Len(arrEtiq(rs_Consult!confnrocol - 2)) = 0 Then
                        arrEtiq(rs_Consult!confnrocol - 2) = rs_Consult!confetiq
                    End If
                Case "AC"
                    If Len(arrAcu(rs_Consult!confnrocol - 2)) = 0 Then
                        arrAcu(rs_Consult!confnrocol - 2) = rs_Consult!confval
                    Else
                        arrAcu(rs_Consult!confnrocol - 2) = arrAcu(rs_Consult!confnrocol - 2) & "," & rs_Consult!confval
                    End If
                    
                    'Etiqueta de la columna
                    If Len(arrEtiq(rs_Consult!confnrocol - 2)) = 0 Then
                        arrEtiq(rs_Consult!confnrocol - 2) = rs_Consult!confetiq
                    End If
            End Select
    End Select
    
    rs_Consult.MoveNext
Loop


Flog.writeline Espacios(Tabulador * 0) & "Configuracion de Conceptos y acumulados a mostrar"
For Ind = 0 To 9

    If ((Len(arrConc(Ind)) <> 0) Or (Len(arrAcu(Ind)) <> 0)) Then
        
        Flog.writeline Espacios(Tabulador * 1) & "Columna: " & Ind & " " & arrEtiq(Ind)
        If (Len(arrConc(Ind)) <> 0) Then
            Flog.writeline Espacios(Tabulador * 2) & "Conceptos: " & arrConc(Ind)
        End If
        If (Len(arrAcu(Ind)) <> 0) Then
            Flog.writeline Espacios(Tabulador * 2) & "Acumuladores: " & arrAcu(Ind)
        End If
        
    End If
    
Next
Flog.writeline


'-------------------------------------------------------------------------------------------------
'Periodo
'-------------------------------------------------------------------------------------------------
Titulo = ""
StrSql = "SELECT pliqnro, pliqdesc, pliqdesde, pliqhasta, pliqmes"
StrSql = StrSql & " FROM periodo"
StrSql = StrSql & " WHERE pliqnro = " & pliqNro
OpenRecordset StrSql, rs_Consult
If Not rs_Consult.EOF Then
    Titulo = Titulo & rs_Consult!pliqdesc
End If


'-------------------------------------------------------------------------------------------------
'Fecha de pago del proceso
'-------------------------------------------------------------------------------------------------
StrSql = "SELECT pronro, prodesc, profecpago, profecini, profecfin"
StrSql = StrSql & " FROM proceso"
StrSql = StrSql & " WHERE pronro = " & proNro
OpenRecordset StrSql, rs_Consult
If Not rs_Consult.EOF Then
    fechaPago = rs_Consult!profecpago
    Titulo = Titulo & " - " & rs_Consult!prodesc
End If


'-------------------------------------------------------------------------------------------------
'Datos de la empresa
'-------------------------------------------------------------------------------------------------
empresaNom = ""
empTernro = 0
StrSql = "SELECT  empnro, empnom, ternro, estrnro"
StrSql = StrSql & " FROM empresa"
StrSql = StrSql & " WHERE empnro = " & Empresa
OpenRecordset StrSql, rs_Consult
If Not rs_Consult.EOF Then
    empresaNom = rs_Consult!empnom
    empTernro = rs_Consult!ternro
    Titulo = Titulo & " - " & empresaNom
Else
    Flog.writeline Espacios(Tabulador * 1) & "ERROR. No se encontro la empresa."
End If
Titulo = Titulo & " - Leg Desde " & legDesde & " Hasta " & legHasta


'-------------------------------------------------------------------------------------------------
'Provincia de la empresa
'-------------------------------------------------------------------------------------------------
empresaProv = ""
domiEmpl = ""
StrSql = " SELECT detdom.calle,detdom.nro,localidad.locdesc,detdom.piso,detdom.oficdepto,provincia.provdesc"
StrSql = StrSql & " FROM  cabdom "
StrSql = StrSql & " INNER JOIN detdom ON detdom.domnro = cabdom.domnro"
StrSql = StrSql & " LEFT JOIN localidad ON detdom.locnro = localidad.locnro"
StrSql = StrSql & " LEFT JOIN provincia ON detdom.provnro = provincia.provnro"
StrSql = StrSql & " WHERE  cabdom.domdefault = -1 AND cabdom.ternro = " & empTernro
OpenRecordset StrSql, rs_Consult
If Not rs_Consult.EOF Then
    domiEmpl = IIf(EsNulo(rs_Consult!calle), "", rs_Consult!calle)
    domiEmpl = domiEmpl & IIf(EsNulo(rs_Consult!nro), "", " Nº " & rs_Consult!nro)
    domiEmpl = domiEmpl & IIf(EsNulo(rs_Consult!piso), "", " " & rs_Consult!piso)
    domiEmpl = domiEmpl & IIf(EsNulo(rs_Consult!oficdepto), "", " " & rs_Consult!oficdepto)
    domiEmpl = domiEmpl & IIf(EsNulo(rs_Consult!locdesc), "", ", comuna de " & rs_Consult!locdesc)
    domiEmpl = domiEmpl & IIf(EsNulo(rs_Consult!provdesc), "", " " & rs_Consult!provdesc)
    empresaProv = IIf(EsNulo(rs_Consult!provdesc), "", rs_Consult!provdesc)
End If

'-------------------------------------------------------------------------------------------------
'RUT de la empresa
'-------------------------------------------------------------------------------------------------
RUTEmpr = ""
StrSql = "SELECT nrodoc"
StrSql = StrSql & " FROM ter_doc"
StrSql = StrSql & " WHERE ter_doc.ternro = " & empTernro
StrSql = StrSql & " AND ter_doc.tidnro = 1"
OpenRecordset StrSql, rs_Consult
If Not rs_Consult.EOF Then
    RUTEmpr = rs_Consult!nrodoc
End If


'-------------------------------------------------------------------------------------------------
'Consulta principal de empleados
'-------------------------------------------------------------------------------------------------
Flog.writeline Espacios(Tabulador * 0) & "Buscando los empleados a procesar"
StrSql = "SELECT DISTINCT empleado.ternro, empleado.terape, empleado.ternom, empleado.terape2, empleado.ternom2, empleado.empleg"
StrSql = StrSql & " FROM Empleado"
StrSql = StrSql & " INNER JOIN cabliq ON cabliq.empleado = empleado.ternro"
StrSql = StrSql & " AND cabliq.pronro = " & proNro
StrSql = StrSql & " AND empleado.empleg >= " & legDesde
StrSql = StrSql & " AND empleado.empleg <= " & legHasta
If Estado <> 2 Then
    StrSql = StrSql & " AND empleado.empest = " & Estado
    If Estado = -1 Then
        Titulo = Titulo & " - Activos"
    Else
        Titulo = Titulo & " - Inactivos"
    End If
    
End If
StrSql = StrSql & " ORDER BY empleado.empleg"
OpenRecordset StrSql, rs_Empleados
    
        
'-------------------------------------------------------------------------------------------------
'seteo de las variables de progreso
'-------------------------------------------------------------------------------------------------
Progreso = 0
CEmpleadosAProc = rs_Empleados.RecordCount
If CEmpleadosAProc = 0 Then
   Flog.writeline ""
   Flog.writeline "NO hay empleados"
   Exit Sub
   CEmpleadosAProc = 1
End If
IncPorc = (99 / CEmpleadosAProc)
Flog.writeline Espacios(Tabulador * 0) & "Cantidad de empleados a procesar " & CEmpleadosAProc
Flog.writeline


'-------------------------------------------------------------------------------------------------
'Ciclo principal
'-------------------------------------------------------------------------------------------------
Orden = 0
Flog.writeline Espacios(Tabulador * 0) & "--------------------------------------------------------"
Flog.writeline Espacios(Tabulador * 0) & "Procesando Empleados"
Flog.writeline Espacios(Tabulador * 0) & "--------------------------------------------------------"
Do While Not rs_Empleados.EOF

    Flog.writeline Espacios(Tabulador * 1) & "Empleado " & rs_Empleados!empleg & " - " & rs_Empleados!terape & " " & rs_Empleados!ternom & " ternro (" & rs_Empleados!ternro & ")"
    
    
    'Apellido
    apeNom = rs_Empleados!terape
    If Not EsNulo(rs_Empleados!terape2) Then apeNom = apeNom & " " & rs_Empleados!terape2
    If Not EsNulo(rs_Empleados!ternom) Then apeNom = apeNom & " " & rs_Empleados!ternom
    If Not EsNulo(rs_Empleados!ternom2) Then apeNom = apeNom & " " & rs_Empleados!ternom2
        
    'RUT DEL EMPLEADO
    RUTEmpl = ""
    StrSql = "SELECT nrodoc"
    StrSql = StrSql & " FROM ter_doc"
    StrSql = StrSql & " WHERE ter_doc.ternro = " & rs_Empleados!ternro
    StrSql = StrSql & " AND ter_doc.tidnro = 1"
    OpenRecordset StrSql, rs_Consult
    If Not rs_Consult.EOF Then
        RUTEmpl = rs_Consult!nrodoc
    End If
    
    'Ultima Fase
    faseDesde = ""
    faseHasta = ""
    causaBaja = ""
    StrSql = "SELECT altfec, bajfec, causa.caudes FROM fases"
    StrSql = StrSql & " LEFT JOIN causa ON causa.caunro = fases.caunro"
    StrSql = StrSql & " WHERE empleado = " & rs_Empleados!ternro
    StrSql = StrSql & " ORDER BY altfec DESC"
    OpenRecordset StrSql, rs_Consult
    If Not rs_Consult.EOF Then
        If Not EsNulo(rs_Consult!altfec) Then faseDesde = rs_Consult!altfec
        If Not EsNulo(rs_Consult!bajfec) Then faseHasta = rs_Consult!bajfec
        If Not EsNulo(rs_Consult!caudes) Then causaBaja = rs_Consult!caudes
    End If
    
    
    'Puesto
    puesto = ""
    StrSql = "SELECT estrdabr"
    StrSql = StrSql & " From his_estructura"
    StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro=his_estructura.estrnro"
    StrSql = StrSql & " AND htetdesde <= " & ConvFecha(fechaPago) & " And (htethasta Is Null Or htethasta >= " & ConvFecha(fechaPago) & ")"
    StrSql = StrSql & " And his_estructura.tenro = 4 And his_estructura.ternro = " & rs_Empleados!ternro
    OpenRecordset StrSql, rs_Consult
    If Not rs_Consult.EOF Then
        puesto = rs_Consult!estrdabr
    End If
                           
    'Representante Legal
    StrSql = "SELECT empterrepleg,terape,terape2,ternom,ternom2"
    StrSql = StrSql & " FROM empresa"
    StrSql = StrSql & " INNER JOIN tercero ON tercero.ternro=empresa.empterrepleg"
    StrSql = StrSql & " WHERE empresa.empnro = " & Empresa
    OpenRecordset StrSql, rs_Consult
    If Not rs_Consult.EOF Then
        repLegal = rs_Consult!terape & " " & rs_Consult!terape2 & " " & rs_Consult!ternom & " " & rs_Consult!ternom2
    End If

              
    '-------------------------------------------------------------------------------------------------
    'Inserto Cabecera
    '-------------------------------------------------------------------------------------------------
    StrSql = "INSERT INTO rep_finiquito_cl"
    StrSql = StrSql & " (bpronro,ternro,orden,Titulo,provincia,fecha,empresa,RUTempre,RUTemple,"
    StrSql = StrSql & " replegal,direccion,puesto,fecdesde,fechasta,caubaja,apenomb,pliqnro,pronro,empleg)"
    StrSql = StrSql & " VALUES("
    StrSql = StrSql & NroProcesoBatch & ","
    StrSql = StrSql & rs_Empleados!ternro & ","
    StrSql = StrSql & Orden & ","
    StrSql = StrSql & "'" & Mid(Titulo, 1, 200) & "',"
    StrSql = StrSql & "'" & Mid(empresaProv, 1, 200) & "',"
    StrSql = StrSql & ConvFecha(fechaPago) & ","
    StrSql = StrSql & "'" & Mid(empresaNom, 1, 200) & "',"
    StrSql = StrSql & "'" & Mid(RUTEmpr, 1, 50) & "',"
    StrSql = StrSql & "'" & Mid(RUTEmpl, 1, 50) & "',"
    StrSql = StrSql & "'" & Mid(repLegal, 1, 200) & "',"
    StrSql = StrSql & "'" & Mid(domiEmpl, 1, 500) & "',"
    StrSql = StrSql & "'" & Mid(puesto, 1, 200) & "',"
    If Not EsNulo(faseDesde) Then
        StrSql = StrSql & ConvFecha(faseDesde) & ","
    Else
        StrSql = StrSql & "NULL,"
    End If
    If Not EsNulo(faseHasta) Then
        StrSql = StrSql & ConvFecha(faseHasta) & ","
    Else
        StrSql = StrSql & "NULL,"
    End If
    StrSql = StrSql & "'" & Mid(causaBaja, 1, 200) & "',"
    StrSql = StrSql & "'" & Mid(apeNom, 1, 250) & "',"
    StrSql = StrSql & pliqNro & ","
    StrSql = StrSql & proNro & ","
    StrSql = StrSql & rs_Empleados!empleg
    StrSql = StrSql & " )"
    objConn.Execute StrSql, , adExecuteNoRecords
              
              
    '-------------------------------------------------------------------------------------------------
    'Busco los conceptos y acum configurados en el confrep
    '-------------------------------------------------------------------------------------------------
    ordenFila = 0
    For Ind = 0 To 9
       
        If ((Len(arrConc(Ind)) <> 0) Or (Len(arrAcu(Ind)) <> 0)) Then
        
            totalConc = 0
            totalAcu = 0
            
            If (Len(arrConc(Ind)) <> 0) Then
                'Busco los conceptos
                totalConc = BuscarConc(proNro, rs_Empleados!ternro, arrConc(Ind))
            End If
            If (Len(arrAcu(Ind)) <> 0) Then
                'busco los acumuladores
                totalAcu = BuscarAcu(proNro, rs_Empleados!ternro, arrAcu(Ind))
            End If
            
            'Inserto detalle
            StrSql = "INSERT INTO rep_finiquito_cl_det"
            StrSql = StrSql & " (bpronro,ternro,descr,valor,orden)"
            StrSql = StrSql & " VALUES("
            StrSql = StrSql & NroProcesoBatch & ","
            StrSql = StrSql & rs_Empleados!ternro & ","
            StrSql = StrSql & "'" & Mid(arrEtiq(Ind), 1, 200) & "',"
            StrSql = StrSql & totalConc + totalAcu & ","
            StrSql = StrSql & ordenFila & ")"
            objConn.Execute StrSql, , adExecuteNoRecords
            
            ordenFila = ordenFila + 1
        
        End If
        
    Next
              
              
    '-------------------------------------------------------------------------------------------------
    'Actualizo el progreso
    '-------------------------------------------------------------------------------------------------
    Orden = Orden + 1
    Progreso = Progreso + IncPorc
    TiempoAcumulado = GetTickCount

    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso
    StrSql = StrSql & ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso))
    StrSql = StrSql & "' WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    'Paso al siguiente Empleado
    rs_Empleados.MoveNext

Loop
            


If rs_Empleados.State = adStateOpen Then rs_Empleados.Close
If rs_Consult.State = adStateOpen Then rs_Consult.Close

Set rs_Empleados = Nothing
Set rs_Consult = Nothing

Exit Sub

E_Finiquito:
    Flog.writeline "=================================================================="
    Flog.writeline "Error: " & Err.Description
    Flog.writeline "Ultimo SQl Ejecutado: " & StrSql
    Flog.writeline "=================================================================="
    'MyRollbackTrans
    'MyBeginTrans
    Progreso = Progreso + IncPorc
    TiempoAcumulado = GetTickCount
    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & _
             ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & _
             "' WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    'MyCommitTrans
    
    HuboError = True
    Flog.writeline " Error: " & Err.Description


End Sub


Public Function BuscarConc(ByVal proNro As Long, ByVal ternro As Long, ByVal listaConc As String) As Double
Dim aux As Double

Dim rs_Conc As New ADODB.Recordset

    aux = 0
    StrSql = "SELECT SUM(detliq.dlimonto) monto"
    StrSql = StrSql & " FROM cabliq"
    StrSql = StrSql & " INNER JOIN proceso ON proceso.pronro = cabliq.pronro"
    StrSql = StrSql & " AND cabliq.pronro = " & proNro
    StrSql = StrSql & " AND cabliq.empleado = " & ternro
    StrSql = StrSql & " INNER JOIN detliq ON cabliq.cliqnro = detliq.cliqnro"
    StrSql = StrSql & " INNER JOIN concepto ON concepto.concnro = detliq.concnro"
    StrSql = StrSql & " WHERE concepto.concnro IN (" & listaConc & ")"
    OpenRecordset StrSql, rs_Conc
    If Not rs_Conc.EOF Then
        'Control nulo 18/11/2008
        If Not EsNulo(rs_Conc!Monto) Then
            aux = rs_Conc!Monto
        End If
    End If
    
    BuscarConc = aux
    
    If rs_Conc.State = adStateOpen Then rs_Conc.Close
    Set rs_Conc = Nothing
End Function


Public Function BuscarAcu(ByVal proNro As Long, ByVal ternro As Long, ByVal listaAcu As String) As Double
Dim aux As Double

Dim rs_Acu As New ADODB.Recordset

    aux = 0
    StrSql = "SELECT SUM(acu_liq.almonto) monto"
    StrSql = StrSql & " FROM cabliq"
    StrSql = StrSql & " INNER JOIN proceso ON proceso.pronro = cabliq.pronro"
    StrSql = StrSql & " AND cabliq.pronro = " & proNro
    StrSql = StrSql & " AND cabliq.empleado = " & ternro
    StrSql = StrSql & " INNER JOIN acu_liq ON cabliq.cliqnro = acu_liq.cliqnro"
    StrSql = StrSql & " AND acu_liq.acunro IN (" & listaAcu & ")"
    OpenRecordset StrSql, rs_Acu
    If Not rs_Acu.EOF Then
        'Control nulo 18/11/2008
        If Not EsNulo(rs_Acu!Monto) Then
            aux = rs_Acu!Monto
        End If
    End If
    
    BuscarAcu = aux
    
    If rs_Acu.State = adStateOpen Then rs_Acu.Close
    Set rs_Acu = Nothing
End Function

