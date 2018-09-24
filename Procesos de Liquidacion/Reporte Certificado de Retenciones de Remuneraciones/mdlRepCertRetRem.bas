Attribute VB_Name = "mdlRepCertRetRem"
Option Explicit

'Const Version = "1.0"
'Const FechaVersion = "03/05/2013"
'Autor = Sebastian Stremel

'Const Version = "1.1"
'Const FechaVersion = "20/05/2013"
'Autor = Sebastian Stremel
'modificacion = Se busca el logo de la empresa. - se agrego la opcion de la accion de ultimo valor a la columna 10 del confrep

'Const Version = "1.2"
'Const FechaVersion = "17/07/2013"
'Gonzalez Nicolás
'modificacion = Se agrega LEFT al insert de la descripcion de la tabla cabecera.

'Const Version = "1.3"
'Const FechaVersion = "12/08/2013"
'Sebastian Stremel
'modificacion = Se cambia campo logo por empterno en la tabla rep_cert_remun_ret

'Const Version = "1.4"
'Const FechaVersion = "20/08/2013"
'modificacion = Se valida que el representante legal no sea nulo - CAS-16441 - H&A - PERU - Cert Ret y REM 5ta Categoria (CAS-15298) [Entrega 2]
                'Sebastian Stremel

'Const Version = "1.5"
'Const FechaVersion = "20/09/2013"
'modificacion = Se corrige el progreso - CAS-16441 - H&A - PERU - Cert Ret y REM 5ta Categoria (CAS-15298) [Entrega 4]
                'Sebastian Stremel

'Const Version = "1.6"
'Const FechaVersion = "15/04/2015"
'modificacion = Se modifica el proceso para buscar los nuevos datos segun el nuevo formato solicitado - CAS-16441 - H&A - PERU - Cert Ret y REM 5ta Categoria (CAS-15298) [Entrega 4]
                'Sebastian Stremel

'Const Version = "1.7"
'Const FechaVersion = "19/08/2015"
'modificacion = Se corrige error en el proceso que se producia al cerrar un objeto que bajo x circunstancia nunca se abria. - CAS-16441 - H&A - PERU - Cert Ret y REM 5ta Categoria - Versión 4 - Cambios Legales [Entrega 7] (CAS-15298)
                'Sebastian Stremel
                
Const Version = "1.8"
Const FechaVersion = "03/02/2016"
'modificacion = Se busca el logo de la empresa - CAS-35308 - RAET - Nacionalización PERU - Mejoras en certificados legales - Certificado de 5ta

Public Sub Main()
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento inicial del Generador de Reporte de Certificados de retenciones de remuneraciones.
' Autor      : Sebastian Stremel
' Fecha      : 15/04/2013
' Ultima Mod.:
' Fecha:
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

    Nombre_Arch = PathFLog & "Cert_Retencion_Remuneraciones" & "-" & NroProcesoBatch & ".log"
    
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
    
    OpenConnection strconexion, objconnProgreso
    If Err.Number <> 0 Then
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
    StrSql = "SELECT * FROM batch_proceso WHERE btprcnro = 390 AND bpronro =" & NroProcesoBatch
    OpenRecordset StrSql, rs_batch_proceso
    
    TiempoInicialProceso = GetTickCount
    
    If Not rs_batch_proceso.EOF Then
        bprcparam = rs_batch_proceso!bprcparam
        rs_batch_proceso.Close
        Set rs_batch_proceso = Nothing
        Call CertiRemuneraciones(NroProcesoBatch, bprcparam)
    Else
        Flog.writeline "no encontró el proceso"
    End If
    
    TiempoFinalProceso = GetTickCount
    Flog.writeline
    Flog.writeline "**********************************************************"
    Flog.writeline " Proceso Finalizado correctamente "
    Flog.writeline
    Flog.writeline "Tiempo del proceso (milisegundos): " & (TiempoFinalProceso - TiempoInicialProceso)
    Flog.writeline
    Flog.writeline "**********************************************************"

    If HuboError Then
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Error' WHERE bpronro = " & NroProcesoBatch
        objconnProgreso.Execute StrSql, , adExecuteNoRecords
    Else
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado' WHERE bpronro = " & NroProcesoBatch
        objconnProgreso.Execute StrSql, , adExecuteNoRecords
    End If
    
    
Fin:
    Flog.Close
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
    MyBeginTrans
        StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcprogreso = 0, bprcestado = 'Error General', bprcpid = " & PID & " WHERE bpronro = " & NroProcesoBatch
        objconnProgreso.Execute StrSql, , adExecuteNoRecords
    MyCommitTrans
End Sub

Public Sub CertiRemuneraciones(ByVal NroProcesoBatch As Long, ByVal bprcparam As String)

Dim rs_consultas As New ADODB.Recordset
Dim objconnDatos As New ADODB.Connection
OpenConnection strconexion, objconnDatos
Dim rs_confrep As New ADODB.Recordset

Dim parametros() As String
Dim lista_procesos As String
Dim Empnro As Long
Dim legdesde As Long
Dim leghasta As Long
Dim pliqdesde As Integer
Dim pliqhasta As Integer
Dim listaProcesos As String
Dim Fecha As String
Dim tituloreporte As String

Dim empnom As String
Dim replegal As Long
Dim empternro  As Long
Dim fechaGen As String
Dim replegalnom As String
Dim Nombre As String
Dim Apellido As String

Dim tipodocEmp As String
Dim nrodocEmp As String
Dim direccionEmp As String

Dim Empleadotipodoc As String
Dim Empleadonrodoc As String
Dim Empleadodireccion As String

Dim TipoDoc As String
Dim NroDoc As String
Dim direccion As String
Dim listaEmpleados As String

Dim Empleado


Dim terempl As Long
Dim empleadoNom As String
Dim nombreEmpl As String
Dim apellidoEmpl As String

Dim cargo As String
Dim TEcargo As Long
Dim j As Integer

Dim esAcumSueldo As Boolean
Dim sueldo As String
Dim esAcumUtil As Boolean
Dim util As String
Dim esAcumComisiones As Boolean
Dim comisiones As String
Dim esAcumGrat As Boolean
Dim Grat As String
Dim esAcumVac As Boolean
Dim vac As String
Dim esAcumFam As Boolean
Dim fam As String
Dim esAcumGan As Boolean
Dim gan As String
Dim esAcumReintegros As Boolean
Dim reintegros As String
Dim esAcumOtrasRem As Boolean
Dim otrasRem As String
Dim esAcumDeducRenta As Boolean
Dim deducRenta As String
Dim esAcumImpRenta As Boolean
Dim impRenta As String
Dim esAcumCredImp As Boolean
Dim impCred As String


Dim valorSueldo As Double
Dim valorUtil As Double
Dim valorComisiones As Double
Dim valorGrat As Double
Dim valorVac As Double
Dim valorFam As Double
Dim valorGan As Double
Dim valorReintegros As Double
Dim valorOtrasRem As Double
Dim valorDeducrenta As Double
Dim valorImpRenta As Double
Dim valorImpCred As Double

Dim nro As Long

Dim prog As Long
Dim porc As Long

Dim Anio As String

Dim logo As String
Dim anchoLogo As Integer
Dim altoLogo As Integer
Dim accion As String


'variables nuevas 2014-04-14
Dim tipoIngresos As String
Dim codIngresos As String
Dim valorIngresos As Double

Dim tipoRentasPropias As String
Dim codRentasPropias As String
Dim valorRentasPropias As Double

Dim tipoRentasOtrasEmp As String
Dim codRentasOtrasEmp As String
Dim valorRentasOtrasEmp As Double

Dim tipoTotalRentasBrutas As String
Dim codTotalRentasBrutas As String
Dim valorTotalRentasBrutas As Double

Dim tipoRenta5ta As String
Dim codRenta5ta As String
Dim valorRenta5ta As Double

Dim tipoRentaImp As String
Dim codRentaImp As String
Dim valorRentaImp As Double

Dim tipoRenta8 As String
Dim codRenta8 As String
Dim valorRenta8 As Double

Dim tipoRenta14 As String
Dim codRenta14 As String
Dim valorRenta14 As Double

Dim tipoRenta17 As String
Dim codRenta17 As String
Dim valorRenta17 As Double

Dim tipoRenta20 As String
Dim codRenta20 As String
Dim valorRenta20 As Double

Dim tipoRenta30 As String
Dim codRenta30 As String
Dim valorRenta30 As Double

Dim tipoImpTotalRet As String
Dim codImpTotalRet As String
Dim valorImpTotalRet As Double

Dim tipoImpRetOtros As String
Dim codImpRetOtros As String
Dim valorImpRetOtros As Double

Dim tipoImpDev As String
Dim codImpDev As String
Dim valorImpDev As Double

Dim tipoCredImp As String
Dim codCredImp As String
Dim valorCredImp As Double

'levanto datos del confrep
StrSql = " SELECT * FROM confrep "
StrSql = StrSql & " WHERE repnro =387"
StrSql = StrSql & " AND conftipo='TE'"
OpenRecordset StrSql, rs_confrep
If Not rs_confrep.EOF Then
    Do While Not rs_confrep.EOF
        Select Case rs_confrep!confnrocol
            Case 20:
                'cargo
                TEcargo = rs_confrep!confval
        End Select
    rs_confrep.MoveNext
    Loop
Else
    Flog.writeline "No se configuraron los tipos de estructura en el confrep"
End If
rs_confrep.Close

'levanto parametros
parametros = Split(bprcparam, "@")

'empresa
Flog.writeline "Empresa: " & parametros(0)
Empnro = parametros(0)

'legajo desde
legdesde = parametros(1)

'legajo hasta
leghasta = parametros(2)

'periodo desde
pliqdesde = parametros(3)

'periodo hasta
pliqhasta = parametros(4)

'lista de proceso
listaProcesos = parametros(5)

'fecha
Fecha = parametros(6)



'BUSCO EL MAXIMO AÑO DE LOS PROCESOS PARA GUARDAR EL AÑO DEL EJERCICIO GRAVABLE (TODOS LOS PROCESOS DEBERIAN SER DEL MISMO AÑO)
StrSql = " SELECT MAX(profecini) anio FROM proceso "
StrSql = StrSql & " WHERE pronro in(" & listaProcesos & ")"
OpenRecordset StrSql, rs_consultas
If Not rs_consultas.EOF Then
    Anio = rs_consultas!Anio
Else
    Flog.writeline " Se produjo un error buscando el maximo año de los procesos"
End If
rs_consultas.Close

Anio = Year(Anio)
'-----------------------------------------
'BUSCO EL NOMBRE DE LA EMPRESA
'-----------------------------------------
If Empnro = "0" Then
    empnom = "Todas"
Else
    StrSql = " SELECT empnom, empterrepleg, ternro FROM empresa "
    StrSql = StrSql & " WHERE empnro=" & Empnro
    OpenRecordset StrSql, rs_consultas
    If Not rs_consultas.EOF Then
        If EsNulo(rs_consultas!empnom) Then
            empnom = ""
        Else
            empnom = rs_consultas!empnom
        End If
        
        If EsNulo(rs_consultas!empterrepleg) Then
            replegal = 0
        Else
            replegal = rs_consultas!empterrepleg
        End If
        
        empternro = rs_consultas!Ternro
        
    Else
        empnom = ""
        replegal = 0
        empternro = 0
    End If
    rs_consultas.Close
End If


tituloreporte = "Empr:" & empnom
tituloreporte = tituloreporte & " - Leg Desde:" & legdesde
tituloreporte = tituloreporte & " - Leg Hasta:" & leghasta
tituloreporte = tituloreporte & " - Per Desde:" & pliqdesde
tituloreporte = tituloreporte & " - Per Hasta:" & pliqhasta


'busco el logo de la empresa
StrSql = " SELECT tipimdire,terimnombre,tipimanchodef,tipimaltodef FROM ter_imag "
StrSql = StrSql & " INNER JOIN tipoimag ON tipoimag.tipimnro=ter_imag.tipimnro "
StrSql = StrSql & " WHERE Ternro =" & empternro
StrSql = StrSql & " AND ter_imag.tipimnro=1 "
OpenRecordset StrSql, rs_consultas
If Not rs_consultas.EOF Then
    logo = rs_consultas!tipimdire & rs_consultas!terimnombre
    anchoLogo = rs_consultas!tipimanchodef
    altoLogo = rs_consultas!tipimaltodef
    Flog.writeline "Se encontro el logo asociado a la empresa"
Else
    logo = ""
    Flog.writeline "No hay logo asociado a la empresa"
End If
rs_consultas.Close
'hasta aca

'INSERTO LA CABECERA DEL REPORTE
StrSql = "INSERT INTO rep_cert_remun_ret "
StrSql = StrSql & " (bpronro, anio, empternro, remretdesc) "
StrSql = StrSql & " VALUES ("
StrSql = StrSql & NroProcesoBatch & ", "
StrSql = StrSql & "'" & Anio & "', "
StrSql = StrSql & empternro & ", "
'StrSql = StrSql & "'" & tituloreporte & "')"
StrSql = StrSql & "'" & Left(tituloreporte, 100) & "')"
objconnDatos.Execute StrSql, , adExecuteNoRecords
'HASTA ACA

'---------------------------------------
'FECHA DE GENERACION
'---------------------------------------
'fechaGen = Format(Now(), "YYYYMMDD")
fechaGen = ConvFecha(Now())


'---------------------------------------
'BUSCO LA LOCALIDAD DE LA EMPRESA
'---------------------------------------


'---------------------------------------
'BUSCO EL NOMBRE DEL REP LEGAL
'---------------------------------------
StrSql = " SELECT ternom,ternom2,terape,terape2 FROM tercero "
StrSql = StrSql & " WHERE ternro=" & replegal
OpenRecordset StrSql, rs_consultas
If Not rs_consultas.EOF Then
    If Not EsNulo(rs_consultas!ternom2) Then
        Nombre = rs_consultas!ternom & " " & rs_consultas!ternom2
    Else
        Nombre = rs_consultas!ternom & " "
    End If
    
    If Not EsNulo(rs_consultas!terape2) Then
        Apellido = rs_consultas!terape & " " & rs_consultas!terape2
    Else
        Apellido = rs_consultas!terape & " "
    End If
    replegalnom = Nombre & " " & Apellido
    Flog.writeline " Se encontraron los datos del representante legal "
Else
    Flog.writeline "No se encontraron los datos del representante legal "
End If
rs_consultas.Close



'---------------------------------------
'BUSCO DATOS DEL REPRESENTANTE LOCAL
StrSql = " SELECT tipodocu.tidsigla tidsigla, via.viadesc , ter_doc.nrodoc nrodoc, calle,nro,sector,torre,piso,oficdepto,manzana,locdesc "
StrSql = StrSql & " FROM tercero "
StrSql = StrSql & " LEFT JOIN ter_doc on ter_doc.ternro = tercero.ternro"
StrSql = StrSql & " LEFT JOIN tipodocu on tipodocu.tidnro = ter_doc.tidnro"
StrSql = StrSql & " LEFT JOIN cabdom on cabdom.ternro = ter_doc.ternro"
StrSql = StrSql & " LEFT JOIN detdom on detdom.domnro = cabdom.domnro"
StrSql = StrSql & " LEFT JOIN localidad on localidad.locnro = detdom.locnro"
StrSql = StrSql & " LEFT JOIN via on via.vianro = detdom.vianro"
StrSql = StrSql & " WHERE tercero.ternro=" & replegal & " order by ter_doc.tidnro asc"
OpenRecordset StrSql, rs_consultas
If Not rs_consultas.EOF Then
    If Not EsNulo(rs_consultas!tidsigla) Then
        TipoDoc = Left(rs_consultas!tidsigla, 3)
    End If
    
    If Not EsNulo(rs_consultas!NroDoc) Then
        NroDoc = rs_consultas!NroDoc
    End If
    
    If Not EsNulo(rs_consultas!viadesc) Then
        direccion = rs_consultas!viadesc
    End If
    
    If Not EsNulo(rs_consultas!calle) Then
        direccion = direccion & " " & rs_consultas!calle
    End If
    
    If Not EsNulo(rs_consultas!nro) Then
        direccion = direccion & " N°" & rs_consultas!nro
    End If
    
    If Not EsNulo(rs_consultas!sector) Then
        direccion = direccion & " Sector " & rs_consultas!sector
    End If

    If Not EsNulo(rs_consultas!torre) Then
        direccion = direccion & " Torre " & rs_consultas!torre
    End If

    If Not EsNulo(rs_consultas!piso) Then
        direccion = direccion & " Piso " & rs_consultas!piso
    End If

    If Not EsNulo(rs_consultas!oficdepto) Then
        direccion = direccion & " Dpto " & rs_consultas!oficdepto
    End If

    If Not EsNulo(rs_consultas!manzana) Then
        direccion = direccion & " Manzana " & rs_consultas!manzana
    End If

    If Not EsNulo(rs_consultas!locdesc) Then
        direccion = direccion & " Localidad " & rs_consultas!locdesc
    End If
Else
    Flog.writeline " No se encontraron los datos del rep legal"
    
End If
rs_consultas.Close



'---------------------------------------
'BUSCO LOS DATOS DE LA EMPRESA
'---------------------------------------
StrSql = " SELECT tipodocu.tidsigla tidsigla, via.viadesc, ter_doc.nrodoc nrodoc, calle,nro,sector,torre,piso,oficdepto,manzana,locdesc "
StrSql = StrSql & " FROM tercero "
StrSql = StrSql & " LEFT JOIN ter_doc on ter_doc.ternro = tercero.ternro"
StrSql = StrSql & " LEFT JOIN tipodocu on tipodocu.tidnro = ter_doc.tidnro"
StrSql = StrSql & " LEFT JOIN cabdom on cabdom.ternro = ter_doc.ternro"
StrSql = StrSql & " LEFT JOIN detdom on detdom.domnro = cabdom.domnro"
StrSql = StrSql & " LEFT JOIN localidad on localidad.locnro = detdom.locnro"
StrSql = StrSql & " LEFT JOIN via on via.vianro = detdom.vianro"
StrSql = StrSql & " WHERE tercero.ternro=" & empternro & " order by ter_doc.tidnro asc"
OpenRecordset StrSql, rs_consultas
If Not rs_consultas.EOF Then
    If Not EsNulo(rs_consultas!tidsigla) Then
        tipodocEmp = Left(rs_consultas!tidsigla, 3)
    End If
    
    If Not EsNulo(rs_consultas!NroDoc) Then
        nrodocEmp = rs_consultas!NroDoc
    End If
    
    If Not EsNulo(rs_consultas!viadesc) Then
        direccionEmp = rs_consultas!viadesc
    End If
    
    If Not EsNulo(rs_consultas!calle) Then
        direccionEmp = direccionEmp & " " & rs_consultas!calle
    End If
    
    If Not EsNulo(rs_consultas!nro) Then
        direccionEmp = direccionEmp & " N°" & rs_consultas!nro
    End If
    
    If Not EsNulo(rs_consultas!sector) Then
        direccionEmp = direccionEmp & " Sector " & rs_consultas!sector
    End If

    If Not EsNulo(rs_consultas!torre) Then
        direccionEmp = direccionEmp & " Torre " & rs_consultas!torre
    End If

    If Not EsNulo(rs_consultas!piso) Then
        direccionEmp = direccionEmp & " Piso " & rs_consultas!piso
    End If

    If Not EsNulo(rs_consultas!oficdepto) Then
        direccionEmp = direccionEmp & " Dpto " & rs_consultas!oficdepto
    End If

    If Not EsNulo(rs_consultas!manzana) Then
        direccionEmp = direccionEmp & " Manzana " & rs_consultas!manzana
    End If

    If Not EsNulo(rs_consultas!locdesc) Then
        direccionEmp = direccionEmp & " " & rs_consultas!locdesc
    End If
Else
    Flog.writeline " No se encontraron los datos del rep legal"
    
End If
rs_consultas.Close


'---------------------------------------
'BUSCO LOS EMPLEADOS
'---------------------------------------
listaEmpleados = "0"
StrSql = " SELECT DISTINCT empleado FROM cabliq "
StrSql = StrSql & " INNER JOIN empleado ON empleado.ternro=cabliq.empleado"
StrSql = StrSql & " WHERE pronro IN (" & listaProcesos & ")"
StrSql = StrSql & " AND empleado.empleg >=" & legdesde
StrSql = StrSql & " AND empleado.empleg <=" & leghasta
OpenRecordset StrSql, rs_consultas
If rs_consultas.EOF Then
    Flog.writeline "No hay empleados para procesar"
    Exit Sub
Else
    Do While Not rs_consultas.EOF
        listaEmpleados = listaEmpleados & ", " & rs_consultas!Empleado
    rs_consultas.MoveNext
    Loop
End If
rs_consultas.Close

'ciclo por cada empleado para buscar los datos
Empleado = Split(listaEmpleados, ",")

If UBound(Empleado) > 0 Then
    porc = CLng(100) / CLng(UBound(Empleado))

Else
    porc = 100
    Flog.writeline "No hay empleados para procesar"
End If
For j = 1 To UBound(Empleado)
    terempl = Empleado(j)

    nombreEmpl = ""
    apellidoEmpl = ""
    Empleadotipodoc = ""
    empleadoNom = ""
    Empleadotipodoc = ""
    Empleadonrodoc = ""
    Empleadodireccion = ""
    '---------------------------------------
    'BUSCO EL NOMBRE DEL EMPLEADO
    '---------------------------------------
    StrSql = " SELECT ternom,ternom2,terape,terape2 FROM tercero "
    StrSql = StrSql & " WHERE ternro=" & terempl
    OpenRecordset StrSql, rs_consultas
    If Not rs_consultas.EOF Then
        If Not EsNulo(rs_consultas!ternom2) Then
            nombreEmpl = rs_consultas!ternom & " " & rs_consultas!ternom2
        Else
            nombreEmpl = rs_consultas!ternom & " "
        End If

        If Not EsNulo(rs_consultas!terape2) Then
            apellidoEmpl = rs_consultas!terape & " " & rs_consultas!terape2
        Else
            apellidoEmpl = rs_consultas!terape & " "
        End If
        empleadoNom = nombreEmpl & " " & apellidoEmpl
        Flog.writeline " Se encontraron los datos del representante legal "
    Else
        Flog.writeline "No se encontraron los datos del representante legal "
    End If
    rs_consultas.Close

    '---------------------------------------
    'BUSCO LA DIRECCION DEL EMPLEADO
    '---------------------------------------
    StrSql = " SELECT tipodocu.tidsigla tidsigla, via.viadesc, ter_doc.nrodoc nrodoc, calle,nro,sector,torre,piso,oficdepto,manzana,locdesc "
    StrSql = StrSql & " FROM tercero "
    StrSql = StrSql & " LEFT JOIN ter_doc on ter_doc.ternro = tercero.ternro"
    StrSql = StrSql & " LEFT JOIN tipodocu on tipodocu.tidnro = ter_doc.tidnro"
    StrSql = StrSql & " LEFT JOIN cabdom on cabdom.ternro = ter_doc.ternro"
    StrSql = StrSql & " LEFT JOIN detdom on detdom.domnro = cabdom.domnro"
    StrSql = StrSql & " LEFT JOIN localidad on localidad.locnro = detdom.locnro"
    StrSql = StrSql & " LEFT JOIN via on via.vianro = detdom.vianro"
    StrSql = StrSql & " WHERE tercero.ternro=" & terempl & " order by ter_doc.tidnro asc"
    OpenRecordset StrSql, rs_consultas
    If Not rs_consultas.EOF Then
        If Not EsNulo(rs_consultas!tidsigla) Then
            Empleadotipodoc = Left(rs_consultas!tidsigla, 3)
        End If

        If Not EsNulo(rs_consultas!NroDoc) Then
            Empleadonrodoc = rs_consultas!NroDoc
        End If

        If Not EsNulo(rs_consultas!viadesc) Then
            Empleadodireccion = rs_consultas!viadesc
        End If

        If Not EsNulo(rs_consultas!calle) Then
            Empleadodireccion = Empleadodireccion & " " & rs_consultas!calle
        End If

        If Not EsNulo(rs_consultas!nro) Then
            Empleadodireccion = Empleadodireccion & " N°" & rs_consultas!nro
        End If

        If Not EsNulo(rs_consultas!sector) Then
            Empleadodireccion = Empleadodireccion & " Sector " & rs_consultas!sector
        End If

        If Not EsNulo(rs_consultas!torre) Then
            Empleadodireccion = Empleadodireccion & " Torre " & rs_consultas!torre
        End If

        If Not EsNulo(rs_consultas!piso) Then
            Empleadodireccion = Empleadodireccion & " Piso " & rs_consultas!piso
        End If

        If Not EsNulo(rs_consultas!oficdepto) Then
            Empleadodireccion = Empleadodireccion & " Dpto " & rs_consultas!oficdepto
        End If

        If Not EsNulo(rs_consultas!manzana) Then
            Empleadodireccion = Empleadodireccion & " Manzana " & rs_consultas!manzana
        End If

        If Not EsNulo(rs_consultas!locdesc) Then
            Empleadodireccion = Empleadodireccion & " " & rs_consultas!locdesc
        End If
    Else
        Flog.writeline " No se encontraron los datos del empleado"

    End If
    rs_consultas.Close

    '------------------------------
    'BUSCO EL CARGO DEL EMPLEADO
    '------------------------------
    StrSql = "SELECT his_estructura.estrnro, estructura.estrdabr"
    StrSql = StrSql & " From empleado "
    StrSql = StrSql & " INNER JOIN his_estructura ON empleado.ternro = his_estructura.ternro AND his_estructura.tenro =" & TEcargo
    StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro"
    StrSql = StrSql & " WHERE (htetdesde <= " & ConvFecha(Fecha) & ") "
    StrSql = StrSql & " AND (htethasta >= " & ConvFecha(Fecha) & " or htethasta is null)"
    StrSql = StrSql & " AND his_estructura.ternro = " & terempl
    OpenRecordset StrSql, rs_consultas
    If Not rs_consultas.EOF Then
        cargo = rs_consultas!estrdabr
    Else
        cargo = ""
        Flog.writeline "No se encontro la estructura cargo para la fecha:" & Fecha
    End If
    
    '-------------------------------------
    'BUSCO LOS CONCEPTOS DE LA LIQUIDACION
    '-------------------------------------
    'levanto datos del confrep
    StrSql = " SELECT * FROM confrep "
    StrSql = StrSql & " WHERE repnro =387"
    StrSql = StrSql & " AND conftipo IN ('CO','AC')"
    StrSql = StrSql & " ORDER BY confnrocol ASC "
    OpenRecordset StrSql, rs_confrep
    If Not rs_confrep.EOF Then
        Do While Not rs_confrep.EOF
            Select Case rs_confrep!confnrocol
                Case 1:
                    'Total de ingresos
                    If rs_confrep!conftipo = "CO" Then
                        tipoIngresos = "CO"
                        codIngresos = rs_confrep!confval2
                    Else
                        tipoIngresos = "AC"
                        codIngresos = rs_confrep!confval
                    End If
                Case 2:
                    'rentas propias
                    If rs_confrep!conftipo = "CO" Then
                        tipoRentasPropias = "CO"
                        codRentasPropias = rs_confrep!confval2
                    Else
                        tipoRentasPropias = "AC"
                        codRentasPropias = rs_confrep!confval
                    End If
                Case 3:
                    'Rentas de otras empresas
                    If rs_confrep!conftipo = "CO" Then
                        tipoRentasOtrasEmp = "CO"
                        codRentasOtrasEmp = rs_confrep!confval2
                    Else
                        tipoRentasOtrasEmp = "AC"
                        codRentasOtrasEmp = rs_confrep!confval
                    End If
                Case 4:
                    'Total Rentas Brutas
                    If rs_confrep!conftipo = "CO" Then
                        tipoTotalRentasBrutas = "CO"
                        codTotalRentasBrutas = rs_confrep!confval2
                    Else
                        tipoTotalRentasBrutas = "AC"
                        codTotalRentasBrutas = rs_confrep!confval
                    End If
                Case 5:
                    'Renta 5ta categoria
                    If rs_confrep!conftipo = "CO" Then
                        tipoRenta5ta = "CO"
                        codRenta5ta = rs_confrep!confval2
                    Else
                        tipoRenta5ta = "AC"
                        codRenta5ta = rs_confrep!confval
                    End If
                Case 6:
                    'Renta Neta Imponible
                    If rs_confrep!conftipo = "CO" Then
                        tipoRentaImp = "CO"
                        codRentaImp = rs_confrep!confval2
                    Else
                        tipoRentaImp = "AC"
                        codRentaImp = rs_confrep!confval
                    End If
                Case 7:
                    'Impuesto a la renta 8%
                    If rs_confrep!conftipo = "CO" Then
                        tipoRenta8 = "CO"
                        codRenta8 = rs_confrep!confval2
                    Else
                        tipoRenta8 = "AC"
                        codRenta8 = rs_confrep!confval
                    End If
                Case 8:
                    'Impuesto a la renta 14%
                    If rs_confrep!conftipo = "CO" Then
                        tipoRenta14 = "CO"
                        codRenta14 = rs_confrep!confval2
                    Else
                        tipoRenta14 = "AC"
                        codRenta14 = rs_confrep!confval
                    End If
                Case 9:
                    'Impuesto a la renta 17%
                    If rs_confrep!conftipo = "CO" Then
                        tipoRenta17 = "CO"
                        codRenta17 = rs_confrep!confval2
                    Else
                        tipoRenta17 = "AC"
                        codRenta17 = rs_confrep!confval
                    End If
                Case 10:
                    'Impuesto a la renta 20%
                    If rs_confrep!conftipo = "CO" Then
                        tipoRenta20 = "CO"
                        codRenta20 = rs_confrep!confval2
                    Else
                        tipoRenta20 = "AC"
                        codRenta20 = rs_confrep!confval
                    End If
                Case 11:
                    'Impuesto a la renta 30%
                    If rs_confrep!conftipo = "CO" Then
                        tipoRenta30 = "CO"
                        codRenta30 = rs_confrep!confval2
                    Else
                        tipoRenta30 = "AC"
                        codRenta30 = rs_confrep!confval
                    End If
                Case 12:
                    'Impuesto total retenido
                    If rs_confrep!conftipo = "CO" Then
                        tipoImpTotalRet = "CO"
                        codImpTotalRet = rs_confrep!confval2
                    Else
                        tipoImpTotalRet = "AC"
                        codImpTotalRet = rs_confrep!confval
                    End If
                Case 13:
                    'Impuesto retenido otras cias
                    If rs_confrep!conftipo = "CO" Then
                        tipoImpRetOtros = "CO"
                        codImpRetOtros = rs_confrep!confval2
                    Else
                        tipoImpRetOtros = "AC"
                        codImpRetOtros = rs_confrep!confval
                    End If
                Case 14:
                    'Importe de devolucion
                    If rs_confrep!conftipo = "CO" Then
                        tipoImpDev = "CO"
                        codImpDev = rs_confrep!confval2
                    Else
                        tipoImpDev = "AC"
                        codImpDev = rs_confrep!confval
                    End If
                Case 15:
                    'Credito contra el impuesto
                    If rs_confrep!conftipo = "CO" Then
                        tipoCredImp = "CO"
                        codCredImp = rs_confrep!confval2
                    Else
                        tipoCredImp = "AC"
                        codCredImp = rs_confrep!confval
                    End If
            End Select
        rs_confrep.MoveNext
        Loop
    Else
        Flog.writeline "No se configuraron los tipos de estructura en el confrep"
    End If
    rs_confrep.Close
    
    
    'busco los ingresos totales
    valorIngresos = obtenerValorCoAc(tipoIngresos, codIngresos, listaProcesos, terempl)
    
    'busco las rentas propias
    valorRentasPropias = obtenerValorCoAc(tipoRentasPropias, codRentasPropias, listaProcesos, terempl)
    
    'busco las rentas de otras empresas
    valorRentasOtrasEmp = obtenerValorCoAc(tipoRentasOtrasEmp, codRentasOtrasEmp, listaProcesos, terempl)
    
    'Total de Rentas Brutas
    valorTotalRentasBrutas = obtenerValorCoAc(tipoTotalRentasBrutas, codTotalRentasBrutas, listaProcesos, terempl)
    
    'Valor 5TA categoria
    valorRenta5ta = obtenerValorCoAc(tipoRenta5ta, codRenta5ta, listaProcesos, terempl)
    
    'Renta neta imponible
    valorRentaImp = obtenerValorCoAc(tipoRentaImp, codRentaImp, listaProcesos, terempl)
    
    'Impuesto Renta 8%
    valorRenta8 = obtenerValorCoAc(tipoRenta8, codRenta8, listaProcesos, terempl)
    
    'Impuesto Renta 14%
    valorRenta14 = obtenerValorCoAc(tipoRenta14, codRenta14, listaProcesos, terempl)

    'Impuesto Renta 17%
    valorRenta17 = obtenerValorCoAc(tipoRenta17, codRenta17, listaProcesos, terempl)

    'Impuesto Renta 20%
    valorRenta20 = obtenerValorCoAc(tipoRenta20, codRenta20, listaProcesos, terempl)

    'Impuesto Renta 30%
    valorRenta30 = obtenerValorCoAc(tipoRenta30, codRenta30, listaProcesos, terempl)
    
    'Impuesto total retenido
    valorImpTotalRet = obtenerValorCoAc(tipoImpTotalRet, codImpTotalRet, listaProcesos, terempl)
    
    'Impuesto Retenido otras cias
    valorImpRetOtros = obtenerValorCoAc(tipoImpRetOtros, codImpRetOtros, listaProcesos, terempl)
    
    'Importe de devolucion
    valorImpDev = obtenerValorCoAc(tipoImpDev, codImpDev, listaProcesos, terempl)
    
    'Credito contra impuestos
    valorCredImp = obtenerValorCoAc(tipoCredImp, codCredImp, listaProcesos, terempl)
    
    'busco el nro remretnro a insertar
    nro = getLastIdentity(objconnDatos, "rep_cert_remun_ret")
    'hasta aca
    
    '-------------------------------------
    'INSERTO LOS DATOS DEL DETALLE
    '-------------------------------------
    StrSql = " INSERT INTO rep_cert_remun_ret_det "
    StrSql = StrSql & "("
    StrSql = StrSql & "remretnro "
    StrSql = StrSql & ", bpronro"
    StrSql = StrSql & ", ternrorep"
    StrSql = StrSql & ", Fecha"
    StrSql = StrSql & ", localidad"
    StrSql = StrSql & ", nomrep"
    StrSql = StrSql & ", tipdocrep"
    StrSql = StrSql & ", docnrorep"
    StrSql = StrSql & ", domrep"
    StrSql = StrSql & ", tipdocempr"
    StrSql = StrSql & ", docnroempr"
    StrSql = StrSql & ", nomempr"
    StrSql = StrSql & ", domempr"
    StrSql = StrSql & ", nomempl"
    StrSql = StrSql & ", tipdocempl"
    StrSql = StrSql & ", docnroempl"
    StrSql = StrSql & ", domempl"
    StrSql = StrSql & ", distempl"
    StrSql = StrSql & ", cargoempl"
    StrSql = StrSql & ", retempl"
    StrSql = StrSql & ", valor1"
    StrSql = StrSql & ", valor2"
    StrSql = StrSql & ", valor3"
    StrSql = StrSql & ", valor4"
    StrSql = StrSql & ", valor5"
    StrSql = StrSql & ", valor6"
    StrSql = StrSql & ", valor7"
    StrSql = StrSql & ", valor8"
    StrSql = StrSql & ", valor9"
    StrSql = StrSql & ", valor10"
    StrSql = StrSql & ", valor11"
    StrSql = StrSql & ", valor12"
    StrSql = StrSql & ", valor13"
    StrSql = StrSql & ", valor14"
    StrSql = StrSql & ", valor15"
    StrSql = StrSql & ", logo"
    StrSql = StrSql & ", altoLogo"
    StrSql = StrSql & ", anchoLogo"
    StrSql = StrSql & ")"
    StrSql = StrSql & " VALUES"
    StrSql = StrSql & " ("
    StrSql = StrSql & nro & ", "
    StrSql = StrSql & NroProcesoBatch & ", "
    StrSql = StrSql & terempl & ", "
    StrSql = StrSql & fechaGen & ", "
    StrSql = StrSql & "'" & Left(direccionEmp, 300) & "', "
    StrSql = StrSql & "'" & Left(replegalnom, 120) & "', "
    StrSql = StrSql & "'" & Left(TipoDoc, 3) & "', "
    StrSql = StrSql & "'" & Left(NroDoc, 15) & "', "
    StrSql = StrSql & "'" & Left(direccion, 300) & "', "
    StrSql = StrSql & "'" & Left(tipodocEmp, 3) & "', "
    StrSql = StrSql & "'" & Left(nrodocEmp, 15) & "', "
    StrSql = StrSql & "'" & Left(empnom, 300) & "', "
    StrSql = StrSql & "'" & Left(direccionEmp, 300) & "', "
    
    StrSql = StrSql & "'" & Left(empleadoNom, 120) & "', "
    StrSql = StrSql & "'" & Left(Empleadotipodoc, 3) & "', "
    StrSql = StrSql & "'" & Left(Empleadonrodoc, 15) & "', "
    StrSql = StrSql & "'" & Left(Empleadodireccion, 300) & "', "
    StrSql = StrSql & "'-', "
    StrSql = StrSql & "'" & Left(cargo, 120) & "', "
    StrSql = StrSql & "0 "
    StrSql = StrSql & ", " & IIf(EsNulo(valorIngresos), 0, valorIngresos)
    StrSql = StrSql & ", " & IIf(EsNulo(valorRentasPropias), 0, valorRentasPropias)
    StrSql = StrSql & ", " & IIf(EsNulo(valorRentasOtrasEmp), 0, valorRentasOtrasEmp)
    StrSql = StrSql & ", " & IIf(EsNulo(valorTotalRentasBrutas), 0, valorTotalRentasBrutas)
    StrSql = StrSql & ", " & IIf(EsNulo(valorRenta5ta), 0, valorRenta5ta)
    StrSql = StrSql & ", " & IIf(EsNulo(valorRentaImp), 0, valorRentaImp)
    StrSql = StrSql & ", " & IIf(EsNulo(valorRenta8), 0, valorRenta8)
    StrSql = StrSql & ", " & IIf(EsNulo(valorRenta14), 0, valorRenta14)
    StrSql = StrSql & ", " & IIf(EsNulo(valorRenta17), 0, valorRenta17)
    StrSql = StrSql & ", " & IIf(EsNulo(valorRenta20), 0, valorRenta20)
    StrSql = StrSql & ", " & IIf(EsNulo(valorRenta30), 0, valorRenta30)
    StrSql = StrSql & ", " & IIf(EsNulo(valorImpTotalRet), 0, valorImpTotalRet)
    StrSql = StrSql & ", " & IIf(EsNulo(valorImpRetOtros), 0, valorImpRetOtros)
    StrSql = StrSql & ", " & IIf(EsNulo(valorImpDev), 0, valorImpDev)
    StrSql = StrSql & ", " & IIf(EsNulo(valorCredImp), 0, valorCredImp)
    StrSql = StrSql & ",'" & Left(logo, 500) & "'"
    StrSql = StrSql & ", " & IIf(EsNulo(altoLogo), 0, altoLogo)
    StrSql = StrSql & ", " & IIf(EsNulo(anchoLogo), 0, anchoLogo)
    StrSql = StrSql & " )"
    objconnDatos.Execute StrSql, , adExecuteNoRecords
    Flog.writeline " Se inserto el detalle correctamente "

    'ACTUALIZO EL PROGRESO POR CADA EMPLEADO
    prog = CDbl(prog) + CDbl(porc)
    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & prog & " WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    'HASTA ACA
Next


End Sub

Public Function obtenerValorCoAc(tipo, codigo, listaProcesos, tercero)

Dim rs_consultas As New ADODB.Recordset



If tipo = "AC" Then
    StrSql = " SELECT sum(almonto) valor FROM cabliq "
    StrSql = StrSql & " INNER JOIN acu_liq ON acu_liq.cliqnro = cabliq.cliqnro "
    StrSql = StrSql & " WHERE cabliq.pronro in(" & listaProcesos & ")"
    StrSql = StrSql & " AND acu_liq.acunro=" & codigo
    StrSql = StrSql & " AND empleado=" & tercero
    OpenRecordset StrSql, rs_consultas
    If Not rs_consultas.EOF Then
        If Not EsNulo(rs_consultas!Valor) Then
            obtenerValorCoAc = rs_consultas!Valor
        Else
            obtenerValorCoAc = 0
        End If
    Else
        obtenerValorCoAc = 0
        Flog.writeline " No esta liquidado el concepto o acumulador: " & codigo
    End If
Else
    If tipo = "CO" Then
        StrSql = " SELECT sum(dlimonto) valor FROM cabliq "
        StrSql = StrSql & " INNER JOIN detliq ON detliq.cliqnro = cabliq.cliqnro "
        StrSql = StrSql & " INNER JOIN concepto ON concepto.concnro = detliq.concnro"
        StrSql = StrSql & " WHERE cabliq.pronro in(" & listaProcesos & ")"
        StrSql = StrSql & " AND concepto.conccod= '" & codigo & "'"
        StrSql = StrSql & " AND empleado=" & tercero
        OpenRecordset StrSql, rs_consultas
        If Not rs_consultas.EOF Then
            If Not EsNulo(rs_consultas!Valor) Then
                obtenerValorCoAc = rs_consultas!Valor
            Else
                obtenerValorCoAc = 0
            End If
        Else
            obtenerValorCoAc = 0
            Flog.writeline " No esta liquidado el concepto o acumulador: " & codigo
        End If
    End If
End If
End Function
