Attribute VB_Name = "MdlRepARTActualizacion"
Global Const Version = "1.01" ' Cesar Stankunas
Global Const FechaModificacion = "04/08/2009"
Global Const UltimaModificacion = ""    'Encriptacion de string connection

Option Explicit

Global IdUser As String
Global Fecha As Date
Global Hora As String

Global Fecha_Inicio_periodo As Date
Global Fecha_Fin_Periodo As Date
Global StrSql2 As String



Public Sub Main()
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento inicial del Generador de Reportes.
' Autor      : FGZ
' Fecha      : 02/03/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim objconnMain As New ADODB.Connection
Dim strCmdLine
Dim Nombre_Arch As String
Dim HuboError As Boolean
Dim rs_batch_proceso As New ADODB.Recordset
Dim PID As String
Dim bprcparam As String
Dim Proceso_FechaDesde As Date
Dim Proceso_FechaHasta As Date
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

    Nombre_Arch = PathFLog & "Reporte_ART_Actualizacion" & "-" & NroProcesoBatch & ".log"
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Nombre_Arch, True)
    
    ' Obtengo el Process ID
    PID = GetCurrentProcessId
    Flog.writeline "PID = " & PID
    
    'Abro la conexion
    On Error Resume Next
    OpenConnection strconexion, objConn
    If Err.Number <> 0 Or Error_Encrypt Then
        Flog.writeline Espacios(Tabulador * 0) & "Problemas en la conexion"
        Exit Sub
    End If
    
    'Cambio el estado del proceso a Procesando y el PID
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcestado = 'Procesando', bprcpid = " & PID & " WHERE bpronro = " & NroProcesoBatch
    objConn.Execute StrSql, , adExecuteNoRecords
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE btprcnro = 41 AND bpronro =" & NroProcesoBatch
    OpenRecordset StrSql, rs_batch_proceso
    
    TiempoInicialProceso = GetTickCount
    
    If Not rs_batch_proceso.EOF Then
        IdUser = rs_batch_proceso!IdUser
        Fecha = rs_batch_proceso!bprcfecha
        Hora = rs_batch_proceso!bprchora
        Proceso_FechaDesde = rs_batch_proceso!bprcfecdesde
        Proceso_FechaHasta = rs_batch_proceso!bprcfechasta
        bprcparam = rs_batch_proceso!bprcparam
        
        rs_batch_proceso.Close
        Set rs_batch_proceso = Nothing
        Call LevantarParamteros(NroProcesoBatch, bprcparam, Proceso_FechaDesde, Proceso_FechaHasta)
    End If
    
    TiempoFinalProceso = GetTickCount
    Flog.writeline "Tiempo del proceso (milisegundos): " & (TiempoFinalProceso - TiempoInicialProceso)
    
    If Not HuboError Then
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado' WHERE bpronro = " & NroProcesoBatch
    Else
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Error' WHERE bpronro = " & NroProcesoBatch
    End If
    objConn.Execute StrSql, , adExecuteNoRecords

    Flog.Close
    objConn.Close

End Sub


Public Sub ConART01(ByVal bpronro As Long, ByVal Empresa As Long, ByVal Agrupado As Boolean, _
    ByVal AgrupaTE1 As Boolean, ByVal Tenro1 As Long, Estrnro1 As Long, _
    ByVal AgrupaTE2 As Boolean, ByVal Tenro2 As Long, Estrnro2 As Long, _
    ByVal AgrupaTE3 As Boolean, ByVal Tenro3 As Long, Estrnro3 As Long, ByVal Filtro As String, _
    ByVal Orden As String, ByVal Fecdde As Date, Fechta As Date)
' --------------------------------------------------------------------------------------------
' Descripcion: Procedimiento de generacion del Reporte de Actualizacion de datos del ART
' Autor      : FGZ
' Fecha      : 02/03/2004
' Ult. Mod   :
' Fecha      :
' --------------------------------------------------------------------------------------------
Dim FechaDesde As Date
Dim FechaHasta As Date

Dim Detalle As String
Dim Proceso As Long
Dim OK As Boolean
Dim Pais As String
Dim Posimp  As String
Dim Cat_dgi As String
Dim Emp_sucursal As String
Dim Apeynom As String
Dim Tipo_movimiento As String
Dim Cod_osocial As String
Dim Par_remu As Integer
Dim Remuneracion As String
Dim Aux_Remuneracion As String
Dim Cuit As String
Dim Cuil As String
Dim Aux_Cuil As String
Dim Aux_Cuit As String
Dim Estructura As String
Dim SucCod As String
Dim Tidsigla As String
Dim Documento As String
Dim Sexo As String
Dim EstadoCivil As String
Dim CP As String
Dim Localidad As String
Dim Direccion As String
Dim Provincia As String
Dim Telefono As String
Dim Nacimiento As String
Dim TerceroEmpresa As Long
Dim Estrnro_Empresa As Long

Dim I As Integer
Dim Arreglo(5) As Single

Dim Ultimo_Empleado As Long

Dim rs_Reporte As New ADODB.Recordset
Dim rs_Empleados As New ADODB.Recordset
Dim rs_Confrep As New ADODB.Recordset
Dim rs_Rep20 As New ADODB.Recordset
Dim rs_Acu_Mes As New ADODB.Recordset
Dim rs_Detdom As New ADODB.Recordset
Dim rs_Estructura As New ADODB.Recordset
Dim rs_Sucursal As New ADODB.Recordset
Dim rs_Localidad As New ADODB.Recordset
Dim rs_Provincia As New ADODB.Recordset
Dim rs_Telefono As New ADODB.Recordset
Dim rs_Cuil As New ADODB.Recordset
Dim rs_Ter_Doc As New ADODB.Recordset
Dim rs_EstadoCivil As New ADODB.Recordset
Dim rs_Nacionalidad As New ADODB.Recordset
Dim rs_Fases As New ADODB.Recordset

Dim fart
Dim Directorio As String
Dim Archivo As String

'Inicializo
Cat_dgi = "2"


'Archivo de exportacion
StrSql = "SELECT sis_dirsalidas FROM sistema WHERE sisnro = 1 "
OpenRecordset StrSql, rs
If Not rs.EOF Then
    Directorio = Trim(rs!sis_dirsalidas)
Else
    'Exit Sub
End If
Archivo = Directorio & "\rotacionart.txt"
Set fs = CreateObject("Scripting.FileSystemObject")
Set fart = fs.CreateTextFile(Archivo, True)


'Busco el CUIT de la Empresa
StrSql = "SELECT * FROM empresa "
StrSql = StrSql & " INNER JOIN ter_doc ON empresa.ternro = ter_doc.ternro "
'StrSql = StrSql & " WHERE ternro =" & rs_Empleados!ternro & " AND tidnro = 10"
StrSql = StrSql & " AND empresa.empnro =" & Empresa
OpenRecordset StrSql, rs_Ter_Doc
If Not rs_Ter_Doc.EOF Then
    Cuit = Left(CStr(rs_Ter_Doc!nrodoc), 13)
    Aux_Cuit = Replace(CStr(Cuit), "-", "")
    TerceroEmpresa = rs_Ter_Doc!ternro
    Estrnro_Empresa = rs_Ter_Doc!estrnro
Else
    Cuit = Space(13)
    Aux_Cuit = Space(12)
    TerceroEmpresa = 0
End If

StrSql = "SELECT * FROM reporte where reporte.repnro = 32"
OpenRecordset StrSql, rs_Reporte
If rs_Reporte.EOF Then
    Flog.writeline "El Reporte Numero 32 no ha sido Configurado"
    Exit Sub
End If

'Configuracion del Reporte
StrSql = "SELECT * FROM confrep WHERE repnro = " & rs_Reporte!repnro
StrSql = StrSql & " AND confnrocol =1"
OpenRecordset StrSql, rs_Confrep
If rs_Confrep.EOF Then
    Flog.writeline "No se encontró la configuración del Reporte"
    Exit Sub
Else
    'Acumulador del Mes
    Par_remu = rs_Confrep!confval
End If

Fecha_Fin_Periodo = Fechta
Fecha_Inicio_periodo = Fecdde

' Comienzo la transaccion
MyBeginTrans

'Depuracion del Temporario
StrSql = "DELETE FROM rep20 "
StrSql = StrSql & " WHERE iduser = '" & IdUser & "'"
objConn.Execute StrSql, , adExecuteNoRecords

'Busco los procesos a evaluar
StrSql = "SELECT DISTINCT empleado.ternro as Tercero, empleado.*, tercero.* FROM empleado "
StrSql = StrSql & " INNER JOIN tercero ON empleado.ternro = tercero.ternro "
StrSql = StrSql & " INNER JOIN his_estructura Empresa ON empresa.ternro = empleado.ternro "
If AgrupaTE1 Then
    StrSql = StrSql & " INNER JOIN his_estructura TE1 ON te1.ternro = empleado.ternro "
End If
If AgrupaTE2 Then
    StrSql = StrSql & " INNER JOIN his_estructura TE2 ON te2.ternro = empleado.ternro "
End If
If AgrupaTE3 Then
    StrSql = StrSql & " INNER JOIN his_estructura TE3 ON te3.ternro = empleado.ternro "
End If
StrSql = StrSql & " WHERE empresa.tenro = 10 AND empresa.estrnro =" & Estrnro_Empresa & " AND " & Filtro
If AgrupaTE1 Then
    StrSql = StrSql & " AND te1.tenro = " & Tenro1 & " AND "
    If Estrnro1 <> 0 Then
        StrSql = StrSql & " te1.estrno = " & Estrnro1 & " AND "
    End If
    StrSql = StrSql & " (te1.htetdesde <= " & ConvFecha(Fecha_Fin_Periodo) & ") AND "
    StrSql = StrSql & " ((" & ConvFecha(Fecha_Fin_Periodo) & " <= te1.htethasta) or (te1.htethasta is null)) "
End If
If AgrupaTE2 Then
    StrSql = StrSql & " AND te2.tenro = " & Tenro2 & " AND "
    If Estrnro2 <> 0 Then
        StrSql = StrSql & " te2.estrno = " & Estrnro2 & " AND "
    End If
    StrSql = StrSql & " (te2.htetdesde <= " & ConvFecha(Fecha_Fin_Periodo) & ") AND "
    StrSql = StrSql & " ((" & ConvFecha(Fecha_Fin_Periodo) & " <= te2.htethasta) or (te2.htethasta is null))  "
End If
If AgrupaTE3 Then
    StrSql = StrSql & " AND te3.tenro = " & Tenro3 & " AND "
    If Estrnro3 <> 0 Then
        StrSql = StrSql & " te3.estrno = " & Estrnro3 & " AND "
    End If
    StrSql = StrSql & " (te3.htetdesde <= " & ConvFecha(Fecha_Fin_Periodo) & ") AND "
    StrSql = StrSql & " ((" & ConvFecha(Fecha_Fin_Periodo) & " <= te3.htethasta) or (te3.htethasta is null))"
End If
StrSql = StrSql & " ORDER BY " & Orden
OpenRecordset StrSql, rs_Empleados

'seteo de las variables de progreso
Progreso = 0
CEmpleadosAProc = rs_Empleados.RecordCount
If CEmpleadosAProc = 0 Then
   CEmpleadosAProc = 1
End If
IncPorc = (100 / CEmpleadosAProc)

Ultimo_Empleado = -1
Do While Not rs_Empleados.EOF
    For I = 1 To 5
        Arreglo(I) = 0
    Next I
    
    'LA CONDICION DE SALIDA ES LO PRIMERO A PROCESAR PORQUE SINO NO HAY QUE HACER NADA
    'Tipo de Movimiento
    StrSql = "SELECT * FROM fases WHERE empleado = " & rs_Empleados!ternro
    'StrSql = StrSql & " AND estado = -1 "
    StrSql = StrSql & " ORDER BY altfec "
    OpenRecordset StrSql, rs_Fases
    If Not rs_Fases.EOF Then
        rs_Fases.MoveLast
        If CBool(rs_Fases!estado) Then 'busco que fases.altfec   este en el rango que viene por parametro
            If rs_Fases!altfec > Fecdde And rs_Fases!altfec < Fechta Then
                Proceso = True
            Else
                Proceso = False
            End If
        Else
            If rs_Fases!bajfec > Fecdde And rs_Fases!bajfec < Fechta Then
                Proceso = True
            Else
                Proceso = False
            End If
        End If
    End If
    
    If Proceso Then
        'SOLO PROCESAR EMPLEADOS CON MOVIMIENTO ENTRE LAS FECHAS DADAS
        
        'cabdom
        StrSql = " SELECT * FROM cabdom " & _
                 " INNER JOIN detdom ON detdom.domnro = cabdom.domnro " & _
                 " WHERE cabdom.ternro = " & rs_Empleados!ternro & " AND " & _
                 " cabdom.tipnro =1"
        If rs_Detdom.State = adStateOpen Then rs_Detdom.Close
        OpenRecordset StrSql, rs_Detdom
        
        'Sucursal
        StrSql = " SELECT * FROM his_estructura " & _
                 "INNER JOIN estructura ON his_estructura.estrnro = estructura.estrnro " & _
                 " WHERE his_estructura.ternro = " & rs_Empleados!ternro & " AND " & _
                 " his_estructura.tenro = 1 AND " & _
                 " (his_estructura.htetdesde <= " & ConvFecha(Fechta) & ") AND " & _
                 " ((" & ConvFecha(Fechta) & " <= his_estructura.htethasta) or (his_estructura.htethasta is null))"
        If rs_Estructura.State = adStateOpen Then rs_Estructura.Close
        OpenRecordset StrSql, rs_Estructura
        If Not rs_Estructura.EOF Then
            Estructura = rs_Estructura!estrdabr
            SucCod = IIf(Not IsNull(rs_Estructura!estrcodext), rs_Estructura!estrcodext, " ")
            StrSql = " SELECT * FROM sucursal " & _
                     " WHERE estrnro =" & rs_Estructura!estrnro
            If rs_Sucursal.State = adStateOpen Then rs_Sucursal.Close
            OpenRecordset StrSql, rs_Sucursal
        Else
            Estructura = ""
            SucCod = ""
        End If
        
        If Not rs_Detdom.EOF Then
        
            'Localidad
            StrSql = " SELECT * FROM localidad WHERE localidad.locnro = " & rs_Detdom!locnro
            If rs_Localidad.State = adStateOpen Then rs_Localidad.Close
            OpenRecordset StrSql, rs_Localidad
            
            'Provincia
            StrSql = " SELECT * FROM provincia WHERE provincia.provnro = " & rs_Detdom!provnro
            If rs_Provincia.State = adStateOpen Then rs_Provincia.Close
            OpenRecordset StrSql, rs_Provincia
            
            'Telefono
            StrSql = " SELECT * FROM telefono WHERE telefono.domnro = " & rs_Detdom!domnro
            If rs_Telefono.State = adStateOpen Then rs_Telefono.Close
            OpenRecordset StrSql, rs_Telefono
        End If
        
        'Obra Social
        StrSql = " SELECT * FROM his_estructura " & _
                 " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro " & _
                 " WHERE his_estructura.ternro = " & rs_Empleados!ternro & " AND " & _
                 " his_estructura.tenro = 17 AND " & _
                 " (his_estructura.htetdesde <= " & ConvFecha(Fechta) & ") AND " & _
                 " ((" & ConvFecha(Fechta) & " <= his_estructura.htethasta) or (his_estructura.htethasta is null))" & _
                 " ORDER BY his_estructura.htetdesde"
        If rs_Estructura.State = adStateOpen Then rs_Estructura.Close
        OpenRecordset StrSql, rs_Estructura
        If Not rs_Estructura.EOF Then
            'Cod_osocial = IIf(Not IsNull(rs_Estructura!estrcodext), Format(rs_Estructura!estrcodext, "000000"), "000000")
            Cod_osocial = Left(rs_Estructura!estrcodext, 6) & Space(IIf((6 - Len(rs_Estructura!estrcodext)) >= 0, 6 - Len(rs_Estructura!estrcodext), 0))
        Else
            Flog.writeline "No se encontro la Obra Social"
            Cod_osocial = "000000"
        End If
        
        StrSql = "SELECT * FROM acu_mes " & _
                 " WHERE acu_mes.acunro = " & Par_remu & _
                 " AND ternro =" & rs_Empleados!ternro & _
                 " AND amanio =" & Year(Fechta) & _
                 " AND ammes =" & Month(Fechta)
        OpenRecordset StrSql, rs_Acu_Mes
        If Not rs_Acu_Mes.EOF Then
            Remuneracion = Right(CStr(Format(rs_Acu_Mes!ammonto, "0000000.00")), 9)
            Aux_Remuneracion = Replace(Remuneracion, ".", "")
        Else
            Aux_Remuneracion = "000000000" 'Format(0, "000000.00")
        End If
        
        'Tipo de Documento
        StrSql = " SELECT * FROM tercero " & _
                 " INNER JOIN ter_doc ON tercero.ternro = ter_doc.ternro AND ter_doc.tidnro <= 4 " & _
                 " WHERE tercero.ternro= " & rs_Empleados!ternro
        OpenRecordset StrSql, rs_Ter_Doc
        If Not rs_Ter_Doc.EOF Then
            Select Case rs_Ter_Doc!tidnro
            Case 1:
                Tidsigla = "01"
            Case 1:
                Tidsigla = "02"
            Case 1:
                Tidsigla = "03"
            Case 1:
                Tidsigla = "04"
            Case 1:
                Tidsigla = "05"
            Case Else
                Tidsigla = "99"
            End Select
        Else
            Flog.writeline "Error al obtener los datos del Tipo de documento"
            Tidsigla = "99"
        End If
        
        'CUIL
        StrSql = " SELECT cuil.nrodoc FROM tercero " & _
                 " INNER JOIN ter_doc cuil ON (tercero.ternro = cuil.ternro AND cuil.tidnro = 10) " & _
                 " WHERE tercero.ternro= " & rs_Empleados!ternro
        OpenRecordset StrSql, rs_Cuil
        If Not rs_Cuil.EOF Then
            Cuil = Left(CStr(rs_Cuil!nrodoc), 13)
            'Cuil = Replace(CStr(Cuil), "-", "") & Space(1)
            Aux_Cuil = Replace(CStr(Cuil), "-", "")
        Else
            Flog.writeline "Error al obtener los datos del cuil"
            Cuil = Space(13)
            Aux_Cuil = Space(12)
        End If
        
        'Estado Civil
        StrSql = " SELECT * FROM estcivil  " & _
                 " INNER JOIN tercero ON tercero.estcivnro = estcivil.estcivnro " & _
                 " WHERE tercero.ternro= " & rs_Empleados!ternro
        OpenRecordset StrSql, rs_EstadoCivil
        
        'Nacionalidad
        StrSql = " SELECT * FROM nacionalidad  " & _
                 " INNER JOIN tercero ON tercero.nacionalnro = nacionalidad.nacionalnro " & _
                 " WHERE tercero.ternro= " & rs_Empleados!ternro
        OpenRecordset StrSql, rs_Nacionalidad
        
        
        'Si no existe el rep20
        StrSql = "SELECT * FROM rep20 "
        StrSql = StrSql & " WHERE ternro = " & rs_Empleados!ternro
        StrSql = StrSql & " AND bpronro = " & bpronro
        OpenRecordset StrSql, rs_Rep20
    
        If rs_Rep20.EOF Then
            'Inserto
            StrSql = "INSERT INTO rep20 (bpronro,empresa,iduser,fecha,hora"
            StrSql = StrSql & ",ternro,legajo,ternom,ternom2,terape,terape2,terfecnac,tersex,terestciv"
            StrSql = StrSql & ",calle,nro,torre,piso,oficdepto,codigopostal,telnro,tidsigla,nrodoc"
            StrSql = StrSql & ",locdesc,provdesc,nacionalidad,sucursal,cuil,contratacion,ingreso,remuneracion"
            For I = 1 To 5
                If Arreglo(I) <> 0 And Not IsNull(Arreglo(I)) Then
                    StrSql = StrSql & ",columna" & CStr(I)
                End If
            Next I
            StrSql = StrSql & ") VALUES ("
            StrSql = StrSql & bpronro & ","
            StrSql = StrSql & Empresa & ","
            StrSql = StrSql & "'" & IdUser & "',"
            StrSql = StrSql & ConvFecha(Fecha) & ","
            StrSql = StrSql & "'" & Hora & "',"
        
            'Tercero y Legajo
            StrSql = StrSql & rs_Empleados!ternro & ","
            StrSql = StrSql & rs_Empleados!empleg & ","
            'Apellidos y Nombres
            StrSql = StrSql & "'" & rs_Empleados!ternom & "'" & ","
            If Not IsNull(rs_Empleados!ternom2) Then
                StrSql = StrSql & "'" & rs_Empleados!ternom2 & "'" & ","
            Else
                StrSql = StrSql & "' '" & ","
            End If
            StrSql = StrSql & "'" & rs_Empleados!terape & "'" & ","
            If Not IsNull(rs_Empleados!terape2) Then
                StrSql = StrSql & "'" & rs_Empleados!terape2 & "'" & ","
            Else
                StrSql = StrSql & "' '" & ","
            End If
            Apeynom = rs_Empleados!ternom & " " & rs_Empleados!terape & Space(IIf(39 - (Len(rs_Empleados!ternom) + Len(rs_Empleados!terape)) >= 0, 39 - (Len(rs_Empleados!ternom) + Len(rs_Empleados!terape)), 0))
            
            'FechaNacimiento
            StrSql = StrSql & ConvFecha(rs_Empleados!terfecnac) & ","
            Nacimiento = Format(Day(rs_Empleados!terfecnac), "00") & Format(Month(rs_Empleados!terfecnac), "00") & Format(Year(rs_Empleados!terfecnac), "0000")
            
            'Sexo y estado civil
            StrSql = StrSql & "'" & IIf(CBool(rs_Empleados!tersex), "M", "F") & "',"
            Sexo = IIf(CBool(rs_Empleados!tersex), "M", "F")
            If Not rs_EstadoCivil.EOF Then
                StrSql = StrSql & "'" & rs_EstadoCivil!estcivdesabr & "',"
                EstadoCivil = Left(rs_EstadoCivil!estcivdesabr, 3) & Space(IIf((3 - Len(rs_EstadoCivil!estcivdesabr)) >= 0, 3 - Len(rs_EstadoCivil!estcivdesabr), 0)) & Space(4)
            Else
                StrSql = StrSql & "' ',"
                EstadoCivil = Space(7)
            End If
            
            'calle,nro,torre,piso,oficdepto,codigopostal
            If Not rs_Detdom.EOF Then
                StrSql = StrSql & "'" & IIf(Not IsNull(rs_Detdom!calle), rs_Detdom!calle, " ") & "'" & ","
                StrSql = StrSql & "'" & IIf(Not IsNull(rs_Detdom!nro), rs_Detdom!nro, " ") & "'" & ","
                StrSql = StrSql & "'" & IIf(Not IsNull(rs_Detdom!torre), rs_Detdom!torre, " ") & "'" & ","
                StrSql = StrSql & "'" & IIf(Not IsNull(rs_Detdom!piso), rs_Detdom!piso, " ") & "'" & ","
                StrSql = StrSql & "'" & IIf(Not IsNull(rs_Detdom!oficdepto), rs_Detdom!oficdepto, " ") & "'" & ","
                
                Direccion = IIf(Not IsNull(rs_Detdom!calle), rs_Detdom!calle & Space(IIf(30 - Len(rs_Detdom!calle) >= 0, 30 - Len(rs_Detdom!calle), 0)), Space(30))
                Direccion = Direccion & IIf(Not IsNull(rs_Detdom!nro), Format(rs_Detdom!nro, "000000"), Space(6))
                Direccion = Direccion & IIf(Not IsNull(rs_Detdom!piso), Format(rs_Detdom!piso, "00"), Space(2))
                Direccion = Direccion & IIf(Not IsNull(rs_Detdom!oficdepto), rs_Detdom!oficdepto & Space(IIf(4 - Len(rs_Detdom!oficdepto) >= 0, 4 - Len(rs_Detdom!oficdepto), 0)), Space(4))
                
                StrSql = StrSql & "'" & IIf(Not IsNull(rs_Detdom!codigopostal), rs_Detdom!codigopostal, " ") & "'" & ","
                CP = IIf(Not IsNull(rs_Detdom!codigopostal), Format(rs_Detdom!codigopostal, "000000000"), Space(9))
            Else
                StrSql = StrSql & "'  '" & ","
                StrSql = StrSql & "'  '" & ","
                StrSql = StrSql & "'  '" & ","
                StrSql = StrSql & "'  '" & ","
                StrSql = StrSql & "'  '" & ","
                StrSql = StrSql & "'  '" & ","
                Direccion = Space(42)
                CP = Space(9)
            End If
            'Telefono
            If Not rs_Telefono.EOF Then
                StrSql = StrSql & "'" & IIf(Not IsNull(rs_Telefono!telnro), Left(rs_Telefono!telnro, 12), " ") & "'" & ","
                Telefono = IIf(Not IsNull(rs_Telefono!telnro), rs_Telefono!telnro & Space(IIf(12 - Len(rs_Telefono!telnro) >= 0, 12 - Len(rs_Telefono!telnro), 0)), Space(12))
            Else
                StrSql = StrSql & "'  '" & ","
                Telefono = Space(12)
            End If
            'Tipo de doc
            StrSql = StrSql & "'" & Tidsigla & "',"
            'Nro de doc
            If Not rs_Ter_Doc.EOF Then
                StrSql = StrSql & "'" & rs_Ter_Doc!nrodoc & "'" & ","
                Documento = Format(rs_Ter_Doc!nrodoc, "00000000")
            Else
                StrSql = StrSql & "'  '" & ","
                Documento = Space(8)
            End If
            
            'Localidad
            If Not rs_Localidad.EOF Then
                StrSql = StrSql & "'" & IIf(Not IsNull(rs_Localidad!locdesc), rs_Localidad!locdesc, " ") & "'" & ","
                Localidad = IIf(Not IsNull(rs_Localidad!locdesc), rs_Localidad!locdesc & Space(IIf(30 - Len(rs_Localidad!locdesc) >= 0, 30 - Len(rs_Localidad!locdesc), 0)), Space(30))
            Else
                StrSql = StrSql & "'  '" & ","
                Localidad = Space(30)
            End If
            'Provincia
            Provincia = "0"
            If Not rs_Provincia.EOF Then
                StrSql = StrSql & "'" & IIf(Not IsNull(rs_Provincia!provcodext), rs_Provincia!provcodext, " ") & "'" & ","
                Provincia = IIf(Not IsNull(rs_Provincia!provcodext), Left(rs_Provincia!provcodext, 1), "0")
            Else
                StrSql = StrSql & "'  '" & ","
                Provincia = "0"
            End If
            'Nacionalidad
            If Not rs_Nacionalidad.EOF Then
                'StrSql = StrSql & "'" & IIf(Not IsNull(rs_Nacionalidad!nacionaldes), rs_Nacionalidad!nacionaldes, " ") & "'" & ","
                StrSql = StrSql & IIf(Not IsNull(rs_Nacionalidad!nacionalnro), rs_Nacionalidad!nacionalnro, 0) & ","
            Else
                StrSql = StrSql & "0" & ","
            End If
            'Sucursal (el codigo externo de la estructura)
            StrSql = StrSql & "'" & SucCod & "'" & ","
            Emp_sucursal = SucCod & Space(IIf(3 - Len(SucCod) >= 0, 3 - Len(SucCod), 0))
            
            'CUIL
            StrSql = StrSql & "'" & Cuil & "'" & ","
            
            'Contratacion , Ingreso
            If Not rs_Fases.EOF Then
                rs_Fases.MoveLast
            End If
            If Not rs_Fases.EOF Then
                If CBool(rs_Fases!estado) Then
                    StrSql = StrSql & "-1" & ","
                    StrSql = StrSql & ConvFecha(rs_Fases!altfec) & ","
                    Tipo_movimiento = "A"
                Else
                    StrSql = StrSql & "0" & ","
                    StrSql = StrSql & ConvFecha(rs_Fases!bajfec) & ","
                    Tipo_movimiento = "B"
                End If
            Else
                StrSql = StrSql & "0" & ",'',"
                Tipo_movimiento = "B"
            End If

            'Remuneracion
            StrSql = StrSql & Remuneracion
            
            For I = 1 To 5
                If Arreglo(I) <> 0 And Not IsNull(Arreglo(I)) Then
                    StrSql = StrSql & "," & Arreglo(I)
                End If
            Next I
            StrSql = StrSql & ")"
            objConn.Execute StrSql, , adExecuteNoRecords
        End If
       
       
    'COLOCAR la variable REMUNERACION Y CUIT con el formato que corresponda, donde corresponda
    fart.writeline Tipo_movimiento & _
                    Apeynom & _
                    Tidsigla & _
                    Documento & _
                    Sexo & _
                    EstadoCivil & _
                    CP & _
                    Localidad & _
                    Direccion & _
                    Provincia & _
                    Cod_osocial & "001" & _
                    Cuil & _
                    Telefono & _
                    Aux_Remuneracion & _
                    Nacimiento & _
                    Emp_sucursal & _
                    Cuit

    End If
    
    'Actualizo el progreso del Proceso
    Progreso = Progreso + IncPorc
    TiempoAcumulado = GetTickCount
    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & _
             ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & _
             "' WHERE bpronro = " & NroProcesoBatch
    objConn.Execute StrSql, , adExecuteNoRecords
            
            
    'Siguiente empleado
    rs_Empleados.MoveNext
Loop

'Fin de la transaccion
MyCommitTrans

'Cierro todo y libero
If rs_Reporte.State = adStateOpen Then rs_Reporte.Close
If rs_Confrep.State = adStateOpen Then rs_Confrep.Close
If rs_Rep20.State = adStateOpen Then rs_Rep20.Close
If rs_Acu_Mes.State = adStateOpen Then rs_Acu_Mes.Close
If rs_Detdom.State = adStateOpen Then rs_Detdom.Close
If rs_Estructura.State = adStateOpen Then rs_Estructura.Close
If rs_Empleados.State = adStateOpen Then rs_Empleados.Close
If rs_Ter_Doc.State = adStateOpen Then rs_Ter_Doc.Close
If rs_Sucursal.State = adStateOpen Then rs_Sucursal.Close
If rs_Localidad.State = adStateOpen Then rs_Localidad.Close
If rs_Provincia.State = adStateOpen Then rs_Provincia.Close
If rs_Telefono.State = adStateOpen Then rs_Telefono.Close
If rs_Cuil.State = adStateOpen Then rs_Cuil.Close
If rs_EstadoCivil.State = adStateOpen Then rs_EstadoCivil.Close
If rs_Nacionalidad.State = adStateOpen Then rs_Nacionalidad.Close
If rs_Fases.State = adStateOpen Then rs_Fases.Close
 
Set rs_Reporte = Nothing
Set rs_Confrep = Nothing
Set rs_Rep20 = Nothing
Set rs_Acu_Mes = Nothing
Set rs_Detdom = Nothing
Set rs_Estructura = Nothing
Set rs_Empleados = Nothing
Set rs_Ter_Doc = Nothing
Set rs_Sucursal = Nothing
Set rs_Localidad = Nothing
Set rs_Provincia = Nothing
Set rs_Telefono = Nothing
Set rs_Cuil = Nothing
Set rs_EstadoCivil = Nothing
Set rs_Nacionalidad = Nothing
Set rs_Fases = Nothing
Exit Sub

CE:
    HuboError = True
    MyRollbackTrans

    'Cierro todo y libero
    If rs_Reporte.State = adStateOpen Then rs_Reporte.Close
    If rs_Confrep.State = adStateOpen Then rs_Confrep.Close
    If rs_Rep20.State = adStateOpen Then rs_Rep20.Close
    If rs_Acu_Mes.State = adStateOpen Then rs_Acu_Mes.Close
    If rs_Detdom.State = adStateOpen Then rs_Detdom.Close
    If rs_Estructura.State = adStateOpen Then rs_Estructura.Close
    If rs_Empleados.State = adStateOpen Then rs_Empleados.Close
    If rs_Ter_Doc.State = adStateOpen Then rs_Ter_Doc.Close
    If rs_Sucursal.State = adStateOpen Then rs_Sucursal.Close
    If rs_Localidad.State = adStateOpen Then rs_Localidad.Close
    If rs_Provincia.State = adStateOpen Then rs_Provincia.Close
    If rs_Telefono.State = adStateOpen Then rs_Telefono.Close
    If rs_Cuil.State = adStateOpen Then rs_Cuil.Close
    If rs_EstadoCivil.State = adStateOpen Then rs_EstadoCivil.Close
    If rs_Nacionalidad.State = adStateOpen Then rs_Nacionalidad.Close
    If rs_Fases.State = adStateOpen Then rs_Fases.Close
     
    Set rs_Reporte = Nothing
    Set rs_Confrep = Nothing
    Set rs_Rep20 = Nothing
    Set rs_Acu_Mes = Nothing
    Set rs_Detdom = Nothing
    Set rs_Estructura = Nothing
    Set rs_Empleados = Nothing
    Set rs_Ter_Doc = Nothing
    Set rs_Sucursal = Nothing
    Set rs_Localidad = Nothing
    Set rs_Provincia = Nothing
    Set rs_Telefono = Nothing
    Set rs_Cuil = Nothing
    Set rs_EstadoCivil = Nothing
    Set rs_Nacionalidad = Nothing
    Set rs_Fases = Nothing

End Sub

Public Sub LevantarParamteros(ByVal bpronro As Long, ByVal parametros As String, ByVal Desde As Date, ByVal Hasta As Date)
' --------------------------------------------------------------------------------------------
' Descripcion: Procedimiento para levantar los parametros pasados en batch_proceso en bprcparam
' Autor      : FGZ
' Fecha      :
' Ult. Mod   :
' Fecha      :
' --------------------------------------------------------------------------------------------
Dim pos1 As Integer
Dim pos2 As Integer

Dim Empresa As Long
Dim Filtro As String
Dim Orden As String

Dim Tenro1 As Long
Dim Tenro2 As Long
Dim Tenro3 As Long

Dim Estrnro1 As Long
Dim Estrnro2 As Long
Dim Estrnro3 As Long

Dim AgrupaTE1 As Boolean
Dim AgrupaTE2 As Boolean
Dim AgrupaTE3 As Boolean

Dim Agrupado As Boolean

'Inicializacion
Agrupado = False
Tenro1 = 0
Tenro2 = 0
Tenro3 = 0
AgrupaTE1 = False
AgrupaTE2 = False
AgrupaTE3 = False

' Levanto cada parametro por separado, el separador de parametros es "."
If Not IsNull(parametros) Then
    If Len(parametros) >= 1 Then
        pos1 = 1
        pos2 = InStr(pos1, parametros, ".") - 1
        Empresa = Mid(parametros, pos1, pos2 - pos1 + 1)
        
        pos1 = pos2 + 2
        pos2 = InStr(pos1, parametros, ".") - 1
        Filtro = Mid(parametros, pos1, pos2 - pos1 + 1)
        
        'A continuacion pueden venir hasta tres niveles de agrupamiento
        ' cero,uno, dos o tres niveles
        pos1 = pos2 + 2
        pos2 = InStr(pos1, parametros, ".") - 1
        If pos2 > 0 Then
            Agrupado = True
            AgrupaTE1 = True
            Tenro1 = Mid(parametros, pos1, pos2 - pos1 + 1)
            If Tenro1 = 0 Then
                Agrupado = False
                AgrupaTE1 = False
            End If
            pos1 = pos2 + 2
            pos2 = InStr(pos1, parametros, ".") - 1
            If Not (pos2 > 0) Then
                pos2 = Len(parametros)
            End If
            Estrnro1 = Mid(parametros, pos1, pos2 - pos1 + 1)

            pos1 = pos2 + 2
            pos2 = InStr(pos1, parametros, ".") - 1
            If pos2 > 0 Then
                AgrupaTE2 = True
                Tenro2 = Mid(parametros, pos1, pos2 - pos1 + 1)
                If Tenro2 = 0 Then
                    AgrupaTE2 = False
                End If
            
                pos1 = pos2 + 2
                pos2 = InStr(pos1, parametros, ".") - 1
                If Not (pos2 > 0) Then
                    pos2 = Len(parametros)
                End If
                Estrnro2 = Mid(parametros, pos1, pos2 - pos1 + 1)
                
                pos1 = pos2 + 2
                pos2 = InStr(pos1, parametros, ".") - 1
                If pos2 > 0 Then
                    AgrupaTE3 = True
                    Tenro3 = Mid(parametros, pos1, pos2 - pos1 + 1)
                    If Tenro3 = 0 Then
                        AgrupaTE3 = False
                    End If
                
                    pos1 = pos2 + 2
                    pos2 = InStr(pos1, parametros, ".") - 1
                    Estrnro3 = Mid(parametros, pos1, pos2 - pos1 + 1)
                    
                    pos1 = pos2 + 2
                End If
            End If
        End If
    End If
End If

pos2 = Len(parametros)
Orden = Mid(parametros, pos1, pos2 - pos1 + 1)

Call ConART01(bpronro, Empresa, Agrupado, AgrupaTE1, Tenro1, Estrnro1, AgrupaTE2, Tenro2, Estrnro2, AgrupaTE3, Tenro3, Estrnro3, Filtro, Orden, Desde, Hasta)

End Sub


Public Function EsElUltimoEmpleado(ByVal rs As ADODB.Recordset, ByVal Anterior As Long) As Boolean
' --------------------------------------------------------------------------------------------
' Descripcion: Procedimiento para saber si es el ultimo empleado de la secuencia
' Autor      : FGZ
' Fecha      :
' Ult. Mod   :
' Fecha      :
' --------------------------------------------------------------------------------------------
    
    rs.MoveNext
    If rs.EOF Then
        EsElUltimoEmpleado = True
    Else
        If rs!Empleado <> Anterior Then
            EsElUltimoEmpleado = True
        Else
            EsElUltimoEmpleado = False
        End If
    End If
    rs.MovePrevious
End Function

