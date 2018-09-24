Attribute VB_Name = "MdlRepCertificadoAnsses"
Option Explicit

'Version: 1.01
' Correccion en el rango del calculo de los anios
'Const Version = 1.01
'Const FechaVersion = "12/09/2005"

'Const Version = 1.02    'Modificacion fechas / fases hiostoricos
'Const FechaVersion = "13/10/2005"

'Const Version = 1.03    'Modificacion sobre la cantidad de fases e historicos de estructuras
'Const FechaVersion = "31/10/2005"

'Const Version = 1.04    'Modificacion sobre la cantidad de fases e historicos de estructuras, permite varias paginas
'Const FechaVersion = "10/11/2005"

'Const Version = 1.05    'Modificacion se saca los puntos al documento(certificante) y se puso que salga el piso en domicilio radicacion
'Const FechaVersion = "11/11/2005"

'Const Version = 1.06    'Si la fecha de generacion del reporte es mayor a la fecha de baja de empleado ==>
'                        ' tomo como fecha hasta la fecha de baja
'Const FechaVersion = "28/11/2005"

'Const Version = 1.07    'Cambio en el formato del domicilio
'Const FechaVersion = "05/12/2005"

'Const Version = 1.08    'Cambio en el calculo de los trabajados
'Const FechaVersion = "12/12/2005"

'Const Version = 1.09    'Cambio en el calculo de los trabajados. Se agregó una columna configurable Opcional en el confep para restar 1 dia en el ultimo dia de cada fase
'Const FechaVersion = "07/03/2006"

'Const Version = 1.1      'Martin Ferraro - Cambio en el formato de fecha de al insertar en base en Aux_Empleado_FechaNacimiento y Aux_Extincion_Fecha
'Const FechaVersion = "17/04/2006"

'Const Version = 1.11     'FGZ - problemas en el armado de las fechas de las tareas entre empresa y fases
'Const FechaVersion = "05/05/2006"

'Const Version = 1.12
'Const FechaVersion = "09/06/2006"
'Fernando Favre - El campo 'Caracter de los servicios' se configura como una estructura para los empleados
'                  Se debe informar en el confrep el Tipo de Estructura. Poner como conftipo 'TES'

'Const Version = 1.13
'Const FechaVersion = "30/06/2006"
'Fernando Favre - Se cambio la forma de calcular las fases
'                 Se agregaron mas log. Calculaba mas si habia licencias.

'Const Version = 1.14
'Const FechaVersion = "21/07/2006"
''Fernando Favre - El campo 'Oficio u Ocupacion' se configura como una estructura para los empleados
''                  Se debe informar en el confrep el Tipo de Estructura. Poner como conftipo 'TEO'

'Const Version = 1.15
'Const FechaVersion = "02/08/2006"
''FGZ - Estaba utilizando como indice para licencia los indices de las fases. Si no coinciden va a andar mal
''      La marca y fecha de extinsion no estaba bien

'Const Version = 1.16
'Const FechaVersion = "03/08/2006"
''FGZ - Debe mostrar solo 5 licencias, no importa cuales (si las ultias o las primeras) pero en el tiempo total se deben contemplar todas


'Const Version = 1.17
'Const FechaVersion = "15/08/2006"
''FGZ -
''        correccion para que cuente el primer dia trabajado
''        If (primerDetalle) And (Aux_Dias <> 0) Then
''            FGZ -15 / 8 / 2006
''            modifiqué la linea, no se por que estaba ni cuando y quien lo agregó pero estaba calculando 1 dia de mas en el primer mes de la primer fase
''            If Day(Aux_Fase_Desde(I)) = 1 And Day(Aux_Fase_Hasta(I)) <> Cantidad_Dias_Mes Then
''                Aux_Dias = Aux_Dias + 1
''            End If
''        End If
''        primerDetalle = False
'
''FGZ -
''       se cambió la funcion que calcula la cantidad total de dias trabajados (Cuadro superior)
''       Call DIF_FECHAS4(Aux_Fase_Desde(I), IIf(Resta_Uno, Aux_Fase_Hasta(I) - 1, Aux_Fase_Hasta(I)), Aux_Fase_Dias(I), Aux_Fase_Meses(I), Aux_Fase_Anios(I))

'Const Version = "1.18"
'Const FechaVersion = "14/05/2007"
'FGZ -
'        Activé manejador de errores global (en el Main())
'        creé una objeto conexion para el progreso del proceso

'Const Version = "1.19"
'Const FechaVersion = "17/01/2008"
'Martin Ferraro - se cambio la funcion DIF_FECHAS4 por DiferenciaFase ya que el primero no realizaba
'correctamente la diferencia porque no tenia en cuenta si el mes actual era de 28,30 o 31 dias
'Tb se cambio la parte donde inserta los detalles cuando acumulaba (el detalle ya existia)

'Global Const Version = 1.2
'Global Const FechaVersion = "31/08/2009"   'Encriptacion de string connection
'Global Const UltimaModificacion = "Manuel Lopez"
'Global Const UltimaModificacion1 = "Encriptacion de string connection"

Global Const Version = 1.21
Global Const FechaVersion = "26/09/2014"
Global Const UltimaModificacion = "Borrelli Facundo"
Global Const UltimaModificacion1 = "Se quita el empleado.* y se indican los campos de la tabla empleado a utilizar, ya que en oracle rompia"

'------------------------------------------------------------
'------------------------------------------------------------
'------------------------------------------------------------
Global IdUser As String
Global Fecha As Date
Global hora As String

Global Aux_Autoriz_Apenom As String
Global Aux_Autoriz_Docu As String
Global Aux_Autoriz_Prov_Emis As String

Global Aux_Certifi_Corresponde As String
Global Aux_Certifi_Doc_Tipo As String
Global Aux_Certifi_Doc_Nro As String
Global Aux_Certifi_Expedida As String



Public Sub Main()
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento inicial del Generador de Reportes.
' Autor      : FGZ
' Fecha      : 17/02/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim objconnMain As New ADODB.Connection
Dim strCmdLine
Dim Nombre_Arch As String
Dim HuboError As Boolean
Dim rs_batch_proceso As New ADODB.Recordset
Dim bprcparam As String
Dim PID As String
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

On Error Resume Next
OpenConnection strconexion, objConn
If Err.Number <> 0 Or Error_Encrypt Then
 Flog.writeline Espacios(Tabulador * 0) & "Problemas en la conexion"
Exit Sub
End If

On Error Resume Next
OpenConnection strconexion, objconnProgreso
If Err.Number <> 0 Or Error_Encrypt Then
 Flog.writeline Espacios(Tabulador * 0) & "Problemas en la conexion"
Exit Sub
End If

On Error GoTo 0
    
    On Error GoTo ME_Main
    
    Nombre_Arch = PathFLog & "Certificados_Anses" & "-" & NroProcesoBatch & ".log"
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Nombre_Arch, True)
    
    ' Obtengo el Process ID
    PID = GetCurrentProcessId
    Flog.writeline "-----------------------------------------------------------------"
    Flog.writeline "Version = " & Version
    Flog.writeline "Fecha   = " & FechaVersion
    Flog.writeline "Modificacion = " & UltimaModificacion
    Flog.writeline "Modificacion = " & UltimaModificacion1
    Flog.writeline "-----------------------------------------------------------------"
    Flog.writeline
    Flog.writeline "PID = " & PID
    
    
    'Cambio el estado del proceso a Procesando y el PID
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcprogreso = 0, bprcfecinicioej = " & ConvFecha(Date) & ", bprcestado = 'Procesando', bprcpid = " & PID & " WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE (btprcnro = 40 ) AND bpronro =" & NroProcesoBatch
    OpenRecordset StrSql, rs_batch_proceso
    
    TiempoInicialProceso = GetTickCount
    
    If Not rs_batch_proceso.EOF Then
        IdUser = rs_batch_proceso!IdUser
        Fecha = rs_batch_proceso!bprcfecha
        hora = rs_batch_proceso!bprchora
        bprcparam = rs_batch_proceso!bprcparam
        
        rs_batch_proceso.Close
        Set rs_batch_proceso = Nothing
        
        Call LevantarParamteros(NroProcesoBatch, bprcparam)
    End If
    
    TiempoFinalProceso = GetTickCount
    Flog.writeline " "
    Flog.writeline "Tiempo del proceso (milisegundos): " & (TiempoFinalProceso - TiempoInicialProceso)
    
    If Not HuboError Then
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado' WHERE bpronro = " & NroProcesoBatch
    Else
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Error' WHERE bpronro = " & NroProcesoBatch
    End If
    objConn.Execute StrSql, , adExecuteNoRecords

Fin:
    Flog.Close
    objConn.Close
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
        
    'si hay alguna transaccion activa la cierra
    MyRollbackTrans
        
    'Actualizo el progreso
    MyBeginTrans
        StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcprogreso = 0, bprcestado = 'Error General', bprcpid = " & PID & " WHERE bpronro = " & NroProcesoBatch
        objconnProgreso.Execute StrSql, , adExecuteNoRecords
    MyCommitTrans
    GoTo Fin:
End Sub

Public Sub Generar_Reporte(ByVal bpronro As Long, ByVal Empresa As Long, ByVal HFecha As Date)
' --------------------------------------------------------------------------------------------
' Descripcion: Procedimiento de generacion del Reporte de Certificado Anses de Servicios
' Autor      : FGZ
' Fecha      : 15/07/2004
' Ult. Mod   : FGZ - 10/11/2005
' Desc       : agrupa en hojas de hasta 5 movimientos de empresas por c/u.
' --------------------------------------------------------------------------------------------
'Variables auxiliares
'Empresa
Dim Aux_Empresa_Certificante As String
Dim Aux_Empresa_Domicilio As String
Dim Aux_Empresa_CodPostal As String
Dim Aux_Empresa_Cuit As String
Dim Aux_Empresa_Actividad As String
Dim Aux_Empresa_Telefono As String
Dim Aux_Empresa_Nro_Inscripcion As String
'Empleado
Dim Aux_Empleado_Cuil As String
Dim Aux_Empleado_Doc As String
Dim Aux_Empleado_Afiliado As String
Dim Aux_Empleado_Apenom As String
Dim Aux_Empleado_FechaNacimiento As String

Dim Aux_Fase(1 To 100) As String
Dim Aux_Fase_Desde(1 To 100) As Date
Dim Aux_Fase_Hasta(1 To 100) As Date
Dim Aux_Fase_Dias(1 To 100) As Long
Dim Aux_Fase_Meses(1 To 100) As Long
Dim Aux_Fase_Anios(1 To 100) As Long

Dim Aux_Total_Tiempo_Dias As Long
Dim Aux_Total_Tiempo_Meses As Long
Dim Aux_Total_Tiempo_Anios As Long
Dim Aux_Extincion As Boolean
Dim Aux_Extincion_Fecha As String
Dim Aux_Extincion_Nro As String
Dim Aux_Puesto As String

Dim Aux_Lic_Desde(1 To 100) As Date
Dim Aux_Lic_Hasta(1 To 100) As Date
Dim Aux_Lic_Dias(1 To 100) As Long
Dim Aux_Lic_Meses(1 To 100) As Long
Dim Aux_Lic_Anios(1 To 100) As Long
Dim Aux_Total_Lic_Dias As Long
Dim Aux_Total_Lic_Meses As Long
Dim Aux_Total_Lic_Anios As Long

Dim Aux_Total_Aniosxperiodo(1 To 100) As Integer

Dim Aux_FD_Calle As String
Dim Aux_FD_Nro As String
Dim Aux_FD_Piso As String
Dim Aux_FD_Dpto As String
Dim Aux_FD_CodPostal As String
Dim Aux_FD_Localidad As String
Dim Aux_FD_Provincia As String
Dim Aux_FD_Telefono As String
Dim Aux_FD_Observaciones As String

'Generales
Dim Aux_Fuente_Doc As String
Dim Aux_Caracter As String
Dim Aux_CI_Nro As String
Dim Aux_Expedido_por As String

Dim Lista_Licencias As String
Dim Acum_Remu As Long
Dim Acum_Sac As Long
Dim TE_Servicios As Long
Dim TE_Puesto As Long
Dim Encontro1 As Boolean
Dim Encontro2 As Boolean
Dim Encontro3 As Boolean
Dim Encontro4 As Boolean
Dim Encontro5 As Boolean
Dim Encontro6 As Boolean
Dim Primera As Boolean

Dim I As Integer
Dim J As Integer
Dim Aux_Indice As Integer
Dim CantidadFases As Integer
Dim Indice_Inicio As Integer
Dim Indice_Fin As Integer
Dim CantidadPaginas As Integer

Dim Continuar As Boolean
Dim Aux_Acum_Remu As Single
Dim Aux_Acum_Sac As Single

Dim MesActual As Integer
Dim AnioActual As Integer

Dim Aux_Meses As Integer
Dim Aux_Dias As Integer
Dim Aux_Horas As Integer
Dim Aux_Licencias As Integer

Dim Aux_RepNro As Long
Dim AuxFecha_inicio As Date
Dim AuxFecha As Date
Dim Aux_Anio As Integer
Dim Ultimo_Anio_Insertado As Integer
Dim Contempla As Boolean

Dim Aux_HFecha As Date
Dim Fecha_Desde As Date
Dim Fecha_Hasta As Date
Dim primerDetalle As Boolean

Dim Resta_Uno_Configurado As Boolean
Dim Resta_Uno As Boolean
Dim Anio_Inicio_Calculo As Integer
Dim Cantidad_Anios_Calculo As Integer
Dim Cantidad_Dias_Mes As Integer

'FGZ - 02/08/2006
Dim UltimaFase As Boolean
Dim UltimaEstructura As Boolean

'FGZ - 03/08/2006
Dim LS_Lic As Long

'Registros
Dim rs_Confrep As New ADODB.Recordset
Dim rs_Acu_Mes As New ADODB.Recordset
Dim rs_Empleados As New ADODB.Recordset
Dim rs_Fases As New ADODB.Recordset
Dim rs_Cuil As New ADODB.Recordset
Dim rs_Doc As New ADODB.Recordset
Dim rs_Cuit As New ADODB.Recordset
Dim rs_Estructura As New ADODB.Recordset
Dim rs_Empresa As New ADODB.Recordset
Dim rs_Zona As New ADODB.Recordset
Dim rs_Provincia As New ADODB.Recordset
Dim rs_Telefono As New ADODB.Recordset
Dim rs_Estr_Cod As New ADODB.Recordset
Dim rs_HisEstructura As New ADODB.Recordset
Dim rs_Licencias As New ADODB.Recordset
Dim rs_rep_PS62 As New ADODB.Recordset
Dim rs_Det As New ADODB.Recordset

On Error GoTo CE
'Configuracion del Reporte
StrSql = "SELECT * FROM reporte INNER JOIN confrep ON  reporte.repnro = confrep.repnro WHERE reporte.repnro = 16"
OpenRecordset StrSql, rs_Confrep
If rs_Confrep.EOF Then
    Flog.writeline "No se encontró la configuración del Reporte 16"
    Exit Sub
End If

Encontro1 = False
Encontro2 = False
Encontro3 = False
Encontro4 = False
Encontro5 = False
Primera = True
Resta_Uno = False
Resta_Uno_Configurado = False
Do While Not rs_Confrep.EOF
    Select Case rs_Confrep!conftipo
    Case "TD":
        If Primera Then
            Primera = False
            Lista_Licencias = CStr(rs_Confrep!confval)
        Else
            Lista_Licencias = Lista_Licencias & "," & rs_Confrep!confval
        End If
        Encontro1 = True
    Case "REM":
        Acum_Remu = rs_Confrep!confval
        Encontro2 = True
    Case "SAC":
        Acum_Sac = rs_Confrep!confval
        Encontro3 = True
    Case "DIA":
        Resta_Uno_Configurado = True
        Encontro4 = True
    Case "TES":
        TE_Servicios = rs_Confrep!confval
        Encontro5 = True
    Case "TEO":
        TE_Puesto = rs_Confrep!confval
        Encontro6 = True
    End Select
    
    rs_Confrep.MoveNext
Loop

If Not Encontro1 Then
    Flog.writeline "  No se encontro definida ninguna Licencia en la configuración. Se debe configurar como de tipo 'TD'."
End If
If Not Encontro2 Then
    Acum_Remu = 0
    Flog.writeline "  No se encontro el Acumulador Remunerativo en la configuración. Se debe configurar como de tipo 'REM' y debe ser mensual."
End If
If Not Encontro3 Then
    Acum_Sac = 0
    Flog.writeline "  No se encontro el Acumulador SAC en la configuración. Se debe configurar como de tipo 'SAC' y debe ser mensual."
End If
If Not Encontro4 Then
    Flog.writeline "  No se resta un dia a cada fase. Se debe configurar como de tipo 'DIA' en la configuración para que en el calculo de diferencia entre 2 fechas, reste uno."
End If
If Not Encontro5 Then
    TE_Servicios = 0
    Flog.writeline "  No se encontro el Tipo Estr. 'Caracter de los servicios' definida en la configuración. Se debe configurar como de tipo 'TES', sino se toma el valor por default 'Comunes'"
End If
If Not Encontro6 Then
    TE_Puesto = 4
    Flog.writeline "  No se encontro el Tipo Estr. 'Oficio u Ocupacion' definida en la configuración. Se debe configurar como de tipo 'TEO', sino se toma el valor por default del tipo de estructura Puesto (4)"
End If


' Comienzo la transaccion
MyBeginTrans

    '-------------------------------------------------------------------------------------------------
    ' Variable que indica la cantidad de años que se calcularan
    '-------------------------------------------------------------------------------------------------
    Cantidad_Anios_Calculo = 11
    
    
    'Cargo los valores fijos
    Aux_Fuente_Doc = "Registros Rubricados"
    Aux_CI_Nro = " "
    Aux_Expedido_por = " "
    
    Aux_FD_Calle = " "
    Aux_FD_Nro = " "
    Aux_FD_Piso = " "
    Aux_FD_Dpto = " "
    Aux_FD_CodPostal = " "
    Aux_FD_Localidad = " "
    Aux_FD_Provincia = " "
    Aux_FD_Telefono = " "
    Aux_FD_Observaciones = " "
    
    'Busco los datos de la empresa
    StrSql = "SELECT * FROM empresa WHERE empresa.empnro = " & Empresa
    OpenRecordset StrSql, rs_Empresa
    
    If Not rs_Empresa.EOF Then
        Aux_Empresa_Certificante = rs_Empresa!empnom
        Aux_Empresa_Actividad = IIf(Not IsNull(rs_Empresa!empactiv), rs_Empresa!empactiv, " ")
    Else
        Flog.writeline "  No se encontró la empresa con empnro = " & Empresa
        Exit Sub
    End If
    
    'Domicilio y codigo postal
    StrSql = " SELECT * FROM detdom " & _
             " INNER JOIN cabdom ON detdom.domnro = cabdom.domnro " & _
             " INNER JOIN localidad ON localidad.locnro = detdom.locnro " & _
             " WHERE cabdom.ternro = " & rs_Empresa!Ternro
    OpenRecordset StrSql, rs_Zona
    If Not rs_Zona.EOF Then
        Aux_Empresa_Domicilio = rs_Zona!calle
        Aux_Empresa_Domicilio = Aux_Empresa_Domicilio & " " & IIf(Not IsNull(rs_Zona!nro), rs_Zona!nro, " ")
        Aux_Empresa_Domicilio = Aux_Empresa_Domicilio & IIf(Not IsNull(rs_Zona!piso), " Piso " & rs_Zona!piso, " ")
        Aux_Empresa_Domicilio = Aux_Empresa_Domicilio & " " & IIf(Not IsNull(rs_Zona!oficdepto), rs_Zona!oficdepto, " ")
        
        Aux_Empresa_CodPostal = IIf(Not IsNull(rs_Zona!codigopostal), rs_Zona!codigopostal, " ")
        
        Aux_FD_Calle = IIf(Not IsNull(rs_Zona!calle), rs_Zona!calle, " ")
        Aux_FD_Nro = IIf(Not IsNull(rs_Zona!nro), rs_Zona!nro, " ")
        Aux_FD_Piso = IIf(Not IsNull(rs_Zona!piso), rs_Zona!piso, " ")
        Aux_FD_Dpto = IIf(Not IsNull(rs_Zona!oficdepto), rs_Zona!oficdepto, " ")
        Aux_FD_CodPostal = IIf(Not IsNull(rs_Zona!codigopostal), rs_Zona!codigopostal, " ")
        Aux_FD_Localidad = IIf(Not IsNull(rs_Zona!locdesc), rs_Zona!locdesc, " ")
    Else
        Flog.writeline "No se encontró el domicilio y la localidad para la empresa."
        Aux_Empresa_Domicilio = " "
        Aux_Empresa_CodPostal = " "
    End If

    'Busco la provincia
    StrSql = " SELECT * FROM detdom " & _
             " INNER JOIN cabdom ON detdom.domnro = cabdom.domnro " & _
             " INNER JOIN provincia ON provincia.provnro = detdom.provnro " & _
             " WHERE cabdom.ternro = " & rs_Empresa!Ternro
    OpenRecordset StrSql, rs_Provincia
    If Not rs_Provincia.EOF Then
        Aux_FD_Provincia = IIf(Not IsNull(rs_Provincia!provcodext), rs_Provincia!provcodext, " ")
    Else
        Flog.writeline "  No se encontró la provincia para la empresa."
    End If

    'CUIT
    StrSql = " SELECT cuit.nrodoc FROM tercero " & _
             " INNER JOIN ter_doc cuit ON (tercero.ternro = cuit.ternro AND cuit.tidnro = 6) " & _
             " WHERE tercero.ternro= " & rs_Empresa!Ternro
    OpenRecordset StrSql, rs_Cuit
    If Not rs_Cuit.EOF Then
        Aux_Empresa_Cuit = Left(CStr(rs_Cuit!NroDoc), 13)
    Else
        Flog.writeline "  No se encontró el cuit para la empresa."
        Aux_Empresa_Cuit = " "
    End If

    'Telefono
    StrSql = " SELECT * FROM telefono WHERE telefono.domnro = " & rs_Zona!domnro & _
             " AND telefono.teldefault = -1 "
    If rs_Telefono.State = adStateOpen Then rs_Telefono.Close
    OpenRecordset StrSql, rs_Telefono
    If Not rs_Telefono.EOF Then
        Aux_Empresa_Telefono = rs_Telefono!telnro
        Aux_FD_Telefono = rs_Telefono!telnro
    Else
        Flog.writeline "  No se encontró el telefono para la empresa."
        Aux_Empresa_Telefono = " "
    End If

    'nro de inscripcion
    StrSql = "SELECT * FROM estr_cod WHERE estrnro =" & rs_Empresa!Estrnro
    StrSql = StrSql & " AND tcodnro = 18"
    OpenRecordset StrSql, rs_Estr_Cod
    If Not rs_Estr_Cod.EOF Then
        Aux_Empresa_Nro_Inscripcion = CStr(rs_Estr_Cod!nrocod)
    Else
        Flog.writeline "  No se encontró el codigo interno 18 para la inscripción para la empresa."
        Aux_Empresa_Nro_Inscripcion = " "
    End If
    
    
Flog.writeline " "
Flog.writeline "Busco los empleados a procesar"
'---------------------------------------------------------------
'Busco los empleados a procesar
'FB - Se quita el empleado.* y se indican los campos de la tabla empleado a utilizar.
StrSql = "SELECT empleado.empleg, empleado.ternro, tercero.terape, tercero.terape2, tercero.ternom, tercero.ternom2, tercero.terfecnac "
'StrSql = "SELECT empleado.*, tercero.terape, tercero.terape2, tercero.ternom, tercero.ternom2, tercero.terfecnac "
StrSql = StrSql & " FROM  empleado "
StrSql = StrSql & " INNER JOIN batch_empleado ON batch_empleado.ternro = empleado.ternro "
StrSql = StrSql & " INNER JOIN tercero ON empleado.ternro = tercero.ternro "
'StrSql = StrSql & " INNER JOIN his_estructura ON his_estructura.ternro = empleado.ternro and his_estructura.tenro = 10 "
'StrSql = StrSql & " INNER JOIN empresa ON empresa.estrnro = his_estructura.estrnro AND empresa.empnro =" & Empresa
StrSql = StrSql & " WHERE batch_empleado.bpronro = " & NroProcesoBatch
Flog.writeline " -- SQL: " & StrSql
OpenRecordset StrSql, rs_Empleados

'Seteo el progreso
Progreso = 0
CEmpleadosAProc = rs_Empleados.RecordCount
If CEmpleadosAProc = 0 Then
    Flog.writeline "No hay empleados que procesar."
    CEmpleadosAProc = 1
End If
IncPorc = (100 / CEmpleadosAProc)


Do While Not rs_Empleados.EOF
    primerDetalle = True
    Flog.writeline "Empleado " & rs_Empleados!empleg
    
    ' Depuracion del Temporario
    StrSql = "SELECT * FROM rep_ps62 "
    StrSql = StrSql & " WHERE empresa = " & Empresa
    StrSql = StrSql & " AND iduser = '" & IdUser & "'"
    StrSql = StrSql & " AND ternro = " & rs_Empleados!Ternro
    'StrSql = StrSql & " AND hfecha = " & ConvFecha(HFecha)
    'FGZ - 28/11/2005
    StrSql = StrSql & " AND fecha = " & ConvFecha(HFecha)
    OpenRecordset StrSql, rs_rep_PS62
    
    If Not rs_rep_PS62.EOF Then
        Flog.writeline "Depuración de datos calculados a la misma fecha de baja. Al " & HFecha
        Do While Not rs_rep_PS62.EOF
            Flog.writeline "   Detalles"
            'Detalles
            StrSql = "DELETE FROM rep_ps62_det "
            StrSql = StrSql & " WHERE empresa = " & Empresa
            StrSql = StrSql & " AND iduser = '" & IdUser & "'"
            StrSql = StrSql & " AND repnro = " & rs_rep_PS62!repnro
            objConn.Execute StrSql, , adExecuteNoRecords
        
            Flog.writeline "   -- SQL: " & StrSql
            
            Flog.writeline "   Encabezados"
            'Encabezados
            StrSql = "DELETE FROM rep_ps62 "
            StrSql = StrSql & " WHERE empresa = " & Empresa
            StrSql = StrSql & " AND iduser = '" & IdUser & "'"
            StrSql = StrSql & " AND ternro = " & rs_Empleados!Ternro
            StrSql = StrSql & " AND fecha = " & ConvFecha(HFecha)
            objConn.Execute StrSql, , adExecuteNoRecords
        
            Flog.writeline "   -- SQL: " & StrSql
            
            rs_rep_PS62.MoveNext
        Loop
    End If
    
    'buscar el documento ppal del empleado, dde. el documento se de sistema (ppal)
    Flog.writeline " "
    Flog.writeline "Buscar el documento principal del empleado."
    StrSql = " SELECT * FROM ter_doc "
    StrSql = StrSql & " INNER JOIN tipodocu ON ter_doc.tidnro = tipodocu.tidnro "
    StrSql = StrSql & " WHERE ternro = " & rs_Empleados!Ternro
    StrSql = StrSql & " AND ter_doc.tidnro <= 4"
    Flog.writeline "  -- SQL: " & StrSql
    OpenRecordset StrSql, rs_Doc
    If Not rs_Doc.EOF Then
        Aux_Empleado_Doc = IIf(IsNull(rs_Doc!tidsigla), "X", rs_Doc!tidsigla) & "-" & IIf(IsNull(rs_Doc!NroDoc), "00000000", rs_Doc!NroDoc)
    Else
        Flog.writeline "  No se encontró ningun documento disponible."
        Aux_Empleado_Doc = ""
    End If
        
    'Buscar el CUIL
    Flog.writeline "Buscar el CUIL"
    StrSql = " SELECT * from ter_doc WHERE ternro = " & rs_Empleados!Ternro
    StrSql = StrSql & " AND tidnro = 10"
    Flog.writeline "  -- SQL: " & StrSql
    OpenRecordset StrSql, rs_Cuil
    If Not rs_Cuil.EOF Then
        Aux_Empleado_Cuil = rs_Cuil!NroDoc
        Aux_Empleado_Afiliado = rs_Cuil!NroDoc
    Else
        Flog.writeline "  No se encontro el CUIL."
        Aux_Empleado_Cuil = ""
        Aux_Empleado_Afiliado = ""
    End If
    
    Aux_Empleado_Apenom = rs_Empleados!terape & " " & IIf(Not IsNull(rs_Empleados!terape2), rs_Empleados!terape2, "") & " " & rs_Empleados!ternom & " " & IIf(Not IsNull(rs_Empleados!ternom2), rs_Empleados!ternom2, "")
    Aux_Empleado_FechaNacimiento = rs_Empleados!terfecnac
        
    'FGZ - Se calcula a la fecha de baja del empleado salvo que sea mayor que la fecha hasta pasada por parametro
    'Busco la ultima fase del empleado y comparo las fechas
    Aux_HFecha = HFecha
    StrSql = "SELECT * FROM fases "
    StrSql = StrSql & " WHERE real = -1 "
    StrSql = StrSql & " AND empleado = " & rs_Empleados!Ternro
    StrSql = StrSql & " AND altfec <= " & ConvFecha(HFecha)
    StrSql = StrSql & " ORDER BY altfec"
    If rs_Fases.State = adStateOpen Then rs_Fases.Close
    OpenRecordset StrSql, rs_Fases
    Flog.writeline "----------------------------------------------------------------"
    Flog.writeline "Calculo de las fases"
    Flog.writeline " -- SQL: " & StrSql
    CantidadFases = 0
    'FGZ - 02/08/2006
    Aux_Extincion = False
    If Not rs_Fases.EOF Then
        rs_Fases.MoveLast
        If Not EsNulo(rs_Fases!bajfec) Then
            'HFecha = IIf(rs_Fases!bajfec < HFecha, rs_Fases!bajfec, HFecha)
            'FGZ - 02/08/2006
            If rs_Fases!bajfec < HFecha Then
                HFecha = rs_Fases!bajfec
                Aux_Extincion = True
                Aux_Extincion_Fecha = rs_Fases!bajfec
            Else
                HFecha = HFecha
            End If
            'FGZ - 02/08/2006
        End If
        rs_Fases.MoveFirst
        
        'Inicializo
        For I = 1 To 100
            Aux_Fase_Dias(I) = 0
            Aux_Fase_Meses(I) = 0
            Aux_Fase_Anios(I) = 0
            Aux_Total_Aniosxperiodo(I) = 0
        Next I
        
        I = 1
        'FGZ - 02/08/2006
        UltimaFase = False
        Do Until rs_Fases.EOF
            'FGZ - 02/08/2006
            UltimaFase = EsUltimoRegistro(rs_Fases)
            Fecha_Desde = rs_Fases!altfec
            If Not EsNulo(rs_Fases!bajfec) Then
                Fecha_Hasta = rs_Fases!bajfec
            Else
                Fecha_Hasta = HFecha
            End If
            
            Flog.writeline "  == Fase desde " & Fecha_Desde & " al " & Fecha_Hasta
                  
            StrSql = " SELECT * FROM his_estructura "
            StrSql = StrSql & " INNER JOIN empresa ON empresa.estrnro = his_estructura.estrnro AND empresa.empnro =" & Empresa
            StrSql = StrSql & " WHERE his_estructura.ternro = " & rs_Empleados!Ternro & " AND his_estructura.tenro = 10 "
            StrSql = StrSql & " AND his_estructura.htetdesde <= " & ConvFecha(Fecha_Hasta)
            StrSql = StrSql & " AND (his_estructura.htethasta >= " & ConvFecha(Fecha_Desde) & " OR his_estructura.htethasta is null) "
            StrSql = StrSql & " ORDER BY his_estructura.htetdesde"
            If rs_HisEstructura.State = adStateOpen Then rs_HisEstructura.Close
            OpenRecordset StrSql, rs_HisEstructura
            
            If rs_HisEstructura.EOF Then
                Flog.writeline "    El Empleado no tiene la empresa seleccionada asignada en dicha fase."
                Flog.writeline "     -- SQL: " & StrSql
                
                Aux_Extincion = False
                Aux_Extincion_Fecha = ""
                Aux_Extincion_Nro = " "
                
                'FGZ - 02/08/2006
                If UltimaFase Then
                    Aux_Extincion = True
                    Aux_Extincion_Fecha = Fecha_Hasta
                    Aux_Extincion_Nro = " "
                End If
            Else
                Flog.writeline "     -- SQL: " & StrSql
                
                'FGZ - 02/08/2006
                UltimaEstructura = False
                Do Until rs_HisEstructura.EOF
                    'FGZ - 02/08/2006
                    UltimaEstructura = EsUltimoRegistro(rs_HisEstructura)
                    
                    Aux_Fase(I) = " "
                    If rs_HisEstructura!htetdesde >= Fecha_Desde Then
                        Aux_Fase_Desde(I) = rs_HisEstructura!htetdesde
                    Else
                        Aux_Fase_Desde(I) = Fecha_Desde
                    End If
                    
                    'FGZ - 02/08/2006
                    'Aux_Extincion = True
                    If Not EsNulo(rs_HisEstructura!htethasta) Then
                        'FGZ - 02/08/2006
                        If UltimaFase And UltimaEstructura Then
                            If rs_HisEstructura!htethasta <= Fecha_Hasta Then
                                Aux_Extincion = True
                                Aux_Extincion_Fecha = rs_HisEstructura!htethasta
                                'Aux_Extincion_Fecha = IIf(rs_HisEstructura!htethasta <= Fecha_Hasta, rs_HisEstructura!htethasta, Fecha_Hasta)
                            End If
                        End If
                        'FGZ - 02/08/2006
                        Aux_Fase_Hasta(I) = IIf(rs_HisEstructura!htethasta <= Fecha_Hasta, rs_HisEstructura!htethasta, Fecha_Hasta)
                    Else
                        'FGZ - 02/08/2006
                        'comente la linea de abajo
                        'Aux_Extincion_Fecha = Fecha_Hasta
                        Aux_Fase_Hasta(I) = Fecha_Hasta
                    End If
                    
                    If Resta_Uno_Configurado Then
                        ' la fecha hasta esta seteado con el fin de la fases y esta configurado Resta_Uno(columna tipo DIA en confrep)
                        ' ==> debo restar un dia al final
                        If Not EsNulo(rs_HisEstructura!htethasta) Then
                            Resta_Uno = True
                        End If
                    End If
                    
                    Aux_Extincion_Nro = " "

                    Flog.writeline "    Fase definida desde " & Aux_Fase_Desde(I) & " al " & Aux_Fase_Hasta(I)
                    
                    CantidadFases = CantidadFases + 1
                        
                    'FGZ - 07/03/2006 - Resta un dia si esta configurado la columna de tipo DIA en el confrep
                    'Call DIF_FECHAS3(Aux_Fase_Desde(I), IIf(Resta_Uno, Aux_Fase_Hasta(I) - 1, Aux_Fase_Hasta(I)), Aux_Fase_Dias(I), Aux_Fase_Meses(I), Aux_Fase_Anios(I))
                        
                    'FGZ - 15/08/2006
                    'Call DIF_FECHAS4(Aux_Fase_Desde(I), IIf(Resta_Uno, Aux_Fase_Hasta(I) - 1, Aux_Fase_Hasta(I)), Aux_Fase_Dias(I), Aux_Fase_Meses(I), Aux_Fase_Anios(I))
                    Call DiferenciaFase(Aux_Fase_Desde(I), IIf(Resta_Uno, Aux_Fase_Hasta(I) - 1, Aux_Fase_Hasta(I)), Aux_Fase_Dias(I), Aux_Fase_Meses(I), Aux_Fase_Anios(I))
                    I = I + 1
                    
                    rs_HisEstructura.MoveNext
                    
                Loop
                rs_HisEstructura.Close
            
            End If
            
            rs_Fases.MoveNext
        Loop
                
        Anio_Inicio_Calculo = Year(Aux_Fase_Hasta(CantidadFases)) - Cantidad_Anios_Calculo
        
    Else
        Flog.writeline "  No se encontraron fases reales anteriores a la fecha " & HFecha
    End If
    rs_Fases.Close
    Flog.writeline "Fin Calculo de las Fases"
    Flog.writeline "----------------------------------------------------------------"
'    Flog.writeline "----------------------------------------------------------------"
        
    'Inicio = 0
    Indice_Inicio = 1
    Indice_Fin = 5
    If Indice_Fin > CantidadFases Then
        Indice_Fin = CantidadFases
    End If
    If (CantidadFases / 5) > Fix(CantidadFases / 5) Then
        CantidadPaginas = Fix(CantidadFases / 5) + 1
    Else
        CantidadPaginas = Fix(CantidadFases / 5)
    End If
    
    
    For J = 1 To CantidadPaginas
        
        Flog.writeline " "
        Flog.writeline " == Calculo de los Valores de los AÑOS que van desde el " & Year(Aux_Fase_Desde(Indice_Inicio)) & " al " & Year(Aux_Fase_Hasta(Indice_Fin))
        
        'Inicializo
        Aux_Total_Lic_Dias = 0
        Aux_Total_Lic_Meses = 0
        Aux_Total_Lic_Anios = 0
        
        For I = 1 To 5
            Aux_Lic_Dias(I) = 0
            Aux_Lic_Meses(I) = 0
            Aux_Lic_Anios(I) = 0
        Next I
        
        'Inicializo El total de dias, meses y años de las tareas
        Aux_Total_Tiempo_Dias = 0
        Aux_Total_Tiempo_Meses = 0
        Aux_Total_Tiempo_Anios = 0
        
        'calculo el total de tiempo
        For I = Indice_Inicio To Indice_Fin 'indice_inicio + 5
            Aux_Total_Tiempo_Dias = Aux_Total_Tiempo_Dias + Aux_Fase_Dias(I)
            Aux_Total_Tiempo_Meses = Aux_Total_Tiempo_Meses + Aux_Fase_Meses(I)
            Aux_Total_Tiempo_Anios = Aux_Total_Tiempo_Anios + Aux_Fase_Anios(I)
            
            If Aux_Total_Tiempo_Dias >= 30 Then
                Aux_Total_Tiempo_Dias = Aux_Total_Tiempo_Dias - 30
                Aux_Total_Tiempo_Meses = Aux_Total_Tiempo_Meses + 1
            End If
            If Aux_Total_Tiempo_Meses >= 12 Then
                Aux_Total_Tiempo_Anios = Aux_Total_Tiempo_Anios + 1
                Aux_Total_Tiempo_Meses = Aux_Total_Tiempo_Meses - 12
            End If
        Next I
        
        'Buscar las licencias en los rangos de las tareas
        Flog.writeline "    Calulo de licencias entre " & Aux_Fase_Desde(Indice_Inicio) & " al " & Aux_Fase_Hasta(Indice_Fin)
        StrSql = "SELECT * FROM emp_lic "
        StrSql = StrSql & " WHERE empleado = " & rs_Empleados!Ternro
        StrSql = StrSql & " AND tdnro in ( " & Lista_Licencias & ") "
        StrSql = StrSql & " AND ( (elfechadesde >=" & ConvFecha(Aux_Fase_Desde(Indice_Inicio))
        StrSql = StrSql & " AND elfechahasta <=" & ConvFecha(Aux_Fase_Hasta(Indice_Inicio)) & ")"
        For I = Indice_Inicio + 1 To Indice_Fin
            If Not IsNull(Aux_Fase(I)) And Aux_Fase(I) = " " Then
                StrSql = StrSql & " OR (elfechadesde >=" & ConvFecha(Aux_Fase_Desde(I))
                StrSql = StrSql & " AND elfechahasta <=" & ConvFecha(Aux_Fase_Hasta(I)) & ")"
            End If
        Next I
        StrSql = StrSql & ")"
        StrSql = StrSql & " ORDER BY elfechadesde "
        
        Flog.writeline "     -- SQL: " & StrSql
        
        OpenRecordset StrSql, rs_Licencias
        
        I = 1
        'FGZ - 03/08/2006
        'Debo mstrar solo 5 licencias pero debo contemplarlas todas a la hora de sumar
        'Do While Not rs_Licencias.EOF And (I <= 5)
        Do While Not rs_Licencias.EOF
            Aux_Lic_Desde(I) = rs_Licencias!elfechadesde
            Aux_Lic_Hasta(I) = rs_Licencias!elfechahasta
                
            'Calcular la antiguedad
            Call DIF_FECHAS3(Aux_Lic_Desde(I), Aux_Lic_Hasta(I), Aux_Lic_Dias(I), Aux_Lic_Meses(I), Aux_Lic_Anios(I))
            I = I + 1
                
            rs_Licencias.MoveNext
        Loop
        
        'FGZ - 03/08/2006
        LS_Lic = I - 1
        
        'If Not rs_Licencias.EOF Then
        'FGZ - 03/08/2006
        If rs_Licencias.RecordCount > 5 Then
            Flog.writeline "      --------------------------------------------------------------------------------------------------------------"
            Flog.writeline "        WARNING. El empleado posee mas de 5 licencias. El reporte soporta hasta 5."
            Flog.writeline "      --------------------------------------------------------------------------------------------------------------"
        End If
        
        'calculo el total de tiempo
        'For I = 1 To 5
        'FGZ - 03/08/2006
        For I = 1 To LS_Lic
            Aux_Total_Lic_Dias = Aux_Total_Lic_Dias + Aux_Lic_Dias(I)
            Aux_Total_Lic_Meses = Aux_Total_Lic_Meses + Aux_Lic_Meses(I)
            Aux_Total_Lic_Anios = Aux_Total_Lic_Anios + Aux_Lic_Anios(I)
            
            If Aux_Total_Lic_Dias >= 30 Then
                Aux_Total_Lic_Dias = Aux_Total_Lic_Dias - 30
                Aux_Total_Lic_Meses = Aux_Total_Lic_Meses + 1
            End If
            If Aux_Total_Lic_Meses >= 12 Then
                Aux_Total_Lic_Anios = Aux_Total_Lic_Anios + 1
                Aux_Total_Lic_Meses = Aux_Total_Lic_Meses - 12
            End If
        Next I
    
    
        'FGZ - 28/11/2005
        'chequeo la fecha de extinsion
        'esta fecha > fecha hasta ==> pongo la fecha hasta
        If Not EsNulo(Aux_Extincion_Fecha) Then
            If Aux_Extincion_Fecha > HFecha Then
                Aux_Extincion_Fecha = HFecha
            End If
        End If
    
        '-----------------------------------------------------------------------------
        'Inserto en Rep_PS62
        Flog.writeline "      Inserto en rep_PS62 "
        
        StrSql = "INSERT INTO rep_PS62 (bpronro,empresa,iduser,fecha,hfecha,hora,"
        StrSql = StrSql & " certificante, Domicilio,codpostal,cuit,nro_inscrip,empactiv,emptelef,fuente_doc,"
        StrSql = StrSql & " cuil,ternro,empleapenom,fechanac,nro_afiliado,ci_nro,expedido_por,empleDocumento,"
        StrSql = StrSql & " extincion,"
        If Not EsNulo(Aux_Extincion_Fecha) Then
            StrSql = StrSql & " extincion_fecha,"
        End If
        StrSql = StrSql & "extincion_nro,"
        'Tareas
        StrSql = StrSql & " tarea1,tar1_desde,tar1_hasta,tar1_dias,tar1_meses,tar1_anios,"
        Aux_Indice = 1
        For I = Indice_Inicio + 1 To Indice_Fin
            Aux_Indice = Aux_Indice + 1
            If Not IsNull(Aux_Fase(I)) And Aux_Fase(I) = " " Then
                StrSql = StrSql & " tarea" & Aux_Indice & ",tar" & Aux_Indice & "_desde,tar" & Aux_Indice & "_hasta,tar" & Aux_Indice & "_dias,tar" & Aux_Indice & "_meses,tar" & Aux_Indice & "_anios,"
            End If
        Next I
        StrSql = StrSql & " total_tiempo_dias,total_tiempo_meses,total_tiempo_anios,"
        'Licencias
        Aux_Indice = 0
        'For i = Indice_Inicio To Indice_Fin
        'FGZ - 02/08/2006
        'cambie la linea de arriba por la de abajo
        For I = Indice_Inicio To 5
            Aux_Indice = Aux_Indice + 1
            If Not IsNull(Aux_Lic_Desde(I)) And Aux_Lic_Desde(I) <> "00:00:00" Then
                StrSql = StrSql & " lic" & Aux_Indice & "_desde,lic" & Aux_Indice & "_hasta,lic" & Aux_Indice & "_ded_dias,lic" & Aux_Indice & "_ded_meses,lic" & Aux_Indice & "_ded_anios,"
            End If
        Next I
        StrSql = StrSql & " total_lic_ded_dias,total_lic_ded_mes,total_lic_ded_anio,"
        StrSql = StrSql & " fd_calle,fd_nro,fd_piso,fd_depto,fd_codpostal,fd_localidad,fd_provincia,fd_telefono,fd_observaciones,"
        StrSql = StrSql & " autoriz_apenom,autoriz_docu,autoriz_prov_emis,"
        StrSql = StrSql & " certifi_correponde,certifi_doc_tipo,certifi_doc_nro,certifi_expedida"
        StrSql = StrSql & " ) VALUES ("
        
        StrSql = StrSql & bpronro & ","
        StrSql = StrSql & Empresa & ","
        StrSql = StrSql & "'" & Left(IdUser, 20) & "',"
        'StrSql = StrSql & ConvFecha(Fecha) & ","
        'FGZ - 28/11/2005
        StrSql = StrSql & ConvFecha(Aux_HFecha) & ","
        StrSql = StrSql & ConvFecha(HFecha) & ","
        StrSql = StrSql & "'" & Left(hora, 10) & "',"
        
        StrSql = StrSql & "'" & Left(Aux_Empresa_Certificante, 60) & "',"
        StrSql = StrSql & "'" & Left(Aux_Empresa_Domicilio, 100) & "',"
        StrSql = StrSql & "'" & Left(Aux_Empresa_CodPostal, 30) & "',"
        StrSql = StrSql & "'" & Left(Aux_Empresa_Cuit, 30) & "',"
        StrSql = StrSql & "'" & Left(Aux_Empresa_Nro_Inscripcion, 30) & "',"
        StrSql = StrSql & "'" & Left(Aux_Empresa_Actividad, 60) & "',"
        StrSql = StrSql & "'" & Left(Aux_Empresa_Telefono, 30) & "',"
        StrSql = StrSql & "'" & Left(Aux_Fuente_Doc, 60) & "',"
        
        StrSql = StrSql & "'" & Left(Aux_Empleado_Cuil, 30) & "',"
        StrSql = StrSql & rs_Empleados!Ternro & ","
        StrSql = StrSql & "'" & Left(Aux_Empleado_Apenom, 60) & "',"
        'StrSql = StrSql & "'" & Aux_Empleado_FechaNacimiento & "',"
        'StrSql = StrSql & "" & ConvFecha(Aux_Empleado_FechaNacimiento) & ","
        StrSql = StrSql & "'" & Format(Aux_Empleado_FechaNacimiento, "yyyy-mm-dd") & "',"
        StrSql = StrSql & "'" & Left(Aux_Empleado_Afiliado, 30) & "',"
        StrSql = StrSql & "'" & Left(Aux_CI_Nro, 30) & "',"
        StrSql = StrSql & "'" & Left(Aux_Expedido_por, 20) & "',"
        StrSql = StrSql & "'" & Left(Aux_Empleado_Doc, 30) & "',"
        
        StrSql = StrSql & CInt(Aux_Extincion) & ","
        'StrSql = StrSql & "'" & Aux_Extincion_Fecha & "',"
        If Not EsNulo(Aux_Extincion_Fecha) Then
            'StrSql = StrSql & ConvFecha(C_Date(Aux_Extincion_Fecha)) & ","
            StrSql = StrSql & "'" & Format(Aux_Extincion_Fecha, "yyyy-mm-dd") & "',"
        End If
        'StrSql = StrSql & "" & IIf(Not EsNulo(Aux_Extincion_Fecha), ConvFecha(CDate(Aux_Extincion_Fecha)), "NULL") & ","
        StrSql = StrSql & "'" & Left(Aux_Extincion_Nro, 10) & "',"
        
        'tareas
        StrSql = StrSql & "'" & Left(Aux_Fase(Indice_Inicio), 100) & "',"
        StrSql = StrSql & "'" & Format(Aux_Fase_Desde(Indice_Inicio), "yyyy-mm-dd") & "',"
        StrSql = StrSql & "'" & Format(Aux_Fase_Hasta(Indice_Inicio), "yyyy-mm-dd") & "',"
        StrSql = StrSql & "" & Left(Aux_Fase_Dias(Indice_Inicio), 2) & ","
        StrSql = StrSql & "" & Left(Aux_Fase_Meses(Indice_Inicio), 2) & ","
        StrSql = StrSql & "" & Left(Aux_Fase_Anios(Indice_Inicio), 2) & ","
        
        For I = Indice_Inicio + 1 To Indice_Fin
            If Not IsNull(Aux_Fase(I)) And Aux_Fase(I) = " " Then
                StrSql = StrSql & "'" & Left(Aux_Fase(I), 100) & "',"
                StrSql = StrSql & "'" & Format(Aux_Fase_Desde(I), "yyyy-mm-dd") & "',"
                StrSql = StrSql & "'" & Format(Aux_Fase_Hasta(I), "yyyy-mm-dd") & "',"
                StrSql = StrSql & "" & Left(Aux_Fase_Dias(I), 2) & ","
                StrSql = StrSql & "" & Left(Aux_Fase_Meses(I), 2) & ","
                StrSql = StrSql & "" & Left(Aux_Fase_Anios(I), 2) & ","
            End If
        Next I
        'Totales
        StrSql = StrSql & "" & Left(Aux_Total_Tiempo_Dias, 2) & ","
        StrSql = StrSql & "" & Left(Aux_Total_Tiempo_Meses, 2) & ","
        StrSql = StrSql & "" & Left(Aux_Total_Tiempo_Anios, 2) & ","
        
        'Licencias
        'For i = Indice_Inicio To Indice_Fin
        'FGZ - 02/08/2006
        'cambie la linea de arriba por la de abajo
        For I = Indice_Inicio To 5
            If Not IsNull(Aux_Lic_Desde(I)) And Aux_Lic_Desde(I) <> "00:00:00" Then
                StrSql = StrSql & "'" & Format(Aux_Lic_Desde(I), "yyyy-mm-dd") & "',"
                StrSql = StrSql & "'" & Format(Aux_Lic_Hasta(I), "yyyy-mm-dd") & "',"
                StrSql = StrSql & "" & Left(Aux_Lic_Dias(I), 2) & ","
                StrSql = StrSql & "" & Left(Aux_Lic_Meses(I), 2) & ","
                StrSql = StrSql & "" & Left(Aux_Lic_Anios(I), 2) & ","
            End If
        Next I
        'Totales
        StrSql = StrSql & "" & Left(Aux_Total_Lic_Dias, 2) & ","
        StrSql = StrSql & "" & Left(Aux_Total_Lic_Meses, 2) & ","
        StrSql = StrSql & "" & Left(Aux_Total_Lic_Anios, 2) & ","
        
        StrSql = StrSql & "'" & Left(Aux_FD_Calle, 60) & "',"
        StrSql = StrSql & "'" & Left(Aux_FD_Nro, 5) & "',"
        StrSql = StrSql & "'" & Left(Aux_FD_Piso, 5) & "',"
        StrSql = StrSql & "'" & Left(Aux_FD_Dpto, 3) & "',"
        StrSql = StrSql & "'" & Left(Aux_FD_CodPostal, 10) & "',"
        StrSql = StrSql & "'" & Left(Aux_FD_Localidad, 60) & "',"
        StrSql = StrSql & "'" & Left$(Aux_FD_Provincia, 2) & "',"
        StrSql = StrSql & "'" & Left(Aux_FD_Telefono, 20) & "',"
        StrSql = StrSql & "'" & Left(Aux_FD_Observaciones, 100) & "',"
        
        StrSql = StrSql & "'" & Left(Aux_Autoriz_Apenom, 60) & "',"
        StrSql = StrSql & "'" & Left(Aux_Autoriz_Docu, 30) & "',"
        StrSql = StrSql & "'" & Left(Aux_Autoriz_Prov_Emis, 3) & "',"
        
        StrSql = StrSql & "'" & Left(Aux_Certifi_Corresponde, 30) & "',"
        StrSql = StrSql & "'" & Left(Aux_Certifi_Doc_Tipo, 3) & "',"
        StrSql = StrSql & "'" & Left(Aux_Certifi_Doc_Nro, 8) & "',"
        StrSql = StrSql & "'" & Left(Aux_Certifi_Expedida, 3) & "'"
        
        StrSql = StrSql & ")"
        
        Flog.writeline "     -- SQL: " & StrSql
        
        objConn.Execute StrSql, , adExecuteNoRecords
        Aux_RepNro = getLastIdentity(objConn, "rep_PS62")
        Flog.writeline "      LastIdentity " & Aux_RepNro
        
        
        'Correccion en el calculo del intervalo
'        Aux_Anio = (11 - (Aux_Total_Tiempo_Anios))
        Aux_Anio = 0
        'Lleno los detalles
        For I = Indice_Inicio To Indice_Fin
            If Not IsNull(Aux_Fase(I)) And Aux_Fase(I) = " " Then
                MesActual = Month(Aux_Fase_Desde(I))
                AnioActual = Year(Aux_Fase_Desde(I))
                
                Continuar = True
                Do While Continuar
                    If AnioActual <= Anio_Inicio_Calculo Then
                        Flog.writeline "        Mes " & MesActual & " del Año " & AnioActual & " no se calcula"
                        'correccion para que NO cuente el primer dia trabajado, ya que esta fase no se mostrara!!
                        primerDetalle = False
                    Else
                        Flog.writeline "        Mes " & MesActual & " del Año " & AnioActual
                        'Remuneracion
                        Aux_Acum_Remu = 0
                        StrSql = " SELECT * FROM acu_mes "
                        StrSql = StrSql & " WHERE ternro =" & rs_Empleados!Ternro
                        StrSql = StrSql & " AND acunro =" & Acum_Remu
                        StrSql = StrSql & " AND amanio =" & AnioActual
                        StrSql = StrSql & " AND ammes =" & MesActual
                        OpenRecordset StrSql, rs_Acu_Mes
                        If Not rs_Acu_Mes.EOF Then
                            Aux_Acum_Remu = rs_Acu_Mes!ammonto
                        End If
                        'SAC
                        Aux_Acum_Sac = 0
                        StrSql = " SELECT * FROM acu_mes "
                        StrSql = StrSql & " WHERE ternro =" & rs_Empleados!Ternro
                        StrSql = StrSql & " AND acunro =" & Acum_Sac
                        StrSql = StrSql & " AND amanio =" & AnioActual
                        StrSql = StrSql & " AND ammes =" & MesActual
                        OpenRecordset StrSql, rs_Acu_Mes
                        If Not rs_Acu_Mes.EOF Then
                            Aux_Acum_Sac = rs_Acu_Mes!ammonto
                        End If
                        
                        Aux_Licencias = 0
                        AuxFecha_inicio = CDate("01/" & MesActual & "/" & AnioActual)
                        If AuxFecha_inicio < Aux_Fase_Desde(I) Then
                            Aux_Licencias = Day(Aux_Fase_Desde(I)) - 1
                            AuxFecha_inicio = Aux_Fase_Desde(I)
                        End If
                        
                        If MesActual <> 12 Then
                            AuxFecha = CDate("01/" & MesActual + 1 & "/" & AnioActual) - 1
                        Else
                            AuxFecha = CDate("31/12/" & AnioActual)
                        End If
                        If AuxFecha > Aux_Fase_Hasta(I) Then
                            AuxFecha = Aux_Fase_Hasta(I)
                            If MesActual = 1 Or MesActual = 3 Or MesActual = 5 Or MesActual = 7 Or MesActual = 8 Or MesActual = 10 Or MesActual = 12 Then
                                Aux_Licencias = Aux_Licencias + (31 - Day(Aux_Fase_Hasta(I)))
                            Else
                                Aux_Licencias = Aux_Licencias + (30 - Day(Aux_Fase_Hasta(I)))
                            End If
                        End If
                                    
                        'Puesto
                        StrSql = " SELECT * FROM his_estructura " & _
                                 " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro " & _
                                 " WHERE his_estructura.ternro = " & rs_Empleados!Ternro & " AND " & _
                                 " his_estructura.tenro = " & TE_Puesto & " AND " & _
                                 " (his_estructura.htetdesde <= " & ConvFecha(AuxFecha) & ") AND " & _
                                 " ((" & ConvFecha(AuxFecha) & " <= his_estructura.htethasta) or (his_estructura.htethasta is null))" & _
                                 " ORDER BY his_estructura.htetdesde"
                        If rs_Estructura.State = adStateOpen Then rs_Estructura.Close
                        OpenRecordset StrSql, rs_Estructura
                        If rs_Estructura.EOF Then
                            Flog.writeline "          NO se encontro el tipo de estuctura " & TE_Puesto
                        Else
                            Aux_Puesto = rs_Estructura!estrdabr
                        End If
                        
                        '---------------------------------------------------
                        'FAF - 08-06-2006
                        'Caracter de los servicios
                        Aux_Caracter = "Comunes"
                        StrSql = " SELECT * FROM his_estructura " & _
                                 " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro " & _
                                 " WHERE his_estructura.ternro = " & rs_Empleados!Ternro & " AND " & _
                                 " his_estructura.tenro = " & TE_Servicios & " AND " & _
                                 " (his_estructura.htetdesde <= " & ConvFecha(AuxFecha) & ") AND " & _
                                 " ((" & ConvFecha(AuxFecha) & " <= his_estructura.htethasta) or (his_estructura.htethasta is null))" & _
                                 " ORDER BY his_estructura.htetdesde"
                        If rs_Estructura.State = adStateOpen Then rs_Estructura.Close
                        OpenRecordset StrSql, rs_Estructura
                        If Not rs_Estructura.EOF Then
                            Aux_Caracter = rs_Estructura!estrdabr
                        End If
                        '---------------------------------------------------
                        
                        'Buscar las licencias en el mes y año actual
                        StrSql = "SELECT * FROM emp_lic "
                        StrSql = StrSql & " WHERE empleado = " & rs_Empleados!Ternro
                        StrSql = StrSql & " AND tdnro in ( " & Lista_Licencias & ") "
                        StrSql = StrSql & " AND (elfechadesde <=" & ConvFecha(AuxFecha)
                        StrSql = StrSql & " AND elfechahasta >=" & ConvFecha(AuxFecha_inicio) & ")"
                        StrSql = StrSql & " ORDER BY elfechadesde "
                        OpenRecordset StrSql, rs_Licencias
                        Flog.writeline "          - Busco las licencias SQL: " & StrSql
                        
                        Do While Not rs_Licencias.EOF
                            
                            Aux_Licencias = Aux_Licencias + CantidadDeDias(rs_Licencias!elfechadesde, rs_Licencias!elfechahasta, AuxFecha_inicio, AuxFecha)
                        
                            rs_Licencias.MoveNext
                        Loop
                        
                        Aux_Horas = 0
                        If Aux_Licencias = 0 Then
                            Aux_Meses = 1
                            Aux_Dias = 0
                        Else
                            Aux_Meses = 0
                            If MesActual = 2 Then 'Febrero
                                If Biciesto(AnioActual) Then
                                    If Aux_Licencias >= 29 Then
                                        Aux_Dias = 0
                                    Else
                                        Aux_Dias = 29 - Aux_Licencias
                                    End If
                                Else
                                    If Aux_Licencias >= 28 Then
                                        Aux_Dias = 0
                                    Else
                                        Aux_Dias = 28 - Aux_Licencias
                                    End If
                                End If
                            Else
                                If MesActual = 1 Or MesActual = 3 Or MesActual = 5 Or MesActual = 7 Or MesActual = 8 Or MesActual = 10 Or MesActual = 12 Then
                                    Cantidad_Dias_Mes = 31
                                Else
                                    Cantidad_Dias_Mes = 30
                                End If
                                
                                If Aux_Licencias > Cantidad_Dias_Mes Then
                                    Aux_Dias = 0
                                Else
                                    Aux_Dias = Cantidad_Dias_Mes - Aux_Licencias
                                End If
                            End If
                        End If
                        
                        'correccion para que cuente el primer dia trabajado
                        If (primerDetalle) And (Aux_Dias <> 0) Then
                            'FGZ - 15/08/2006
                            'modifiqué la linea, no se por que estaba ni cuando y quien lo agregó pero estaba calculando 1 dia de mas en el primer mes de la primer fase
                            If Day(Aux_Fase_Desde(I)) = 1 And Day(Aux_Fase_Hasta(I)) <> Cantidad_Dias_Mes Then
                                Aux_Dias = Aux_Dias + 1
                            End If
                        End If
                        primerDetalle = False
                        
                        'If (Aux_Meses = 0) And (Aux_Dias = 0) Then
                        '    Aux_Dias = 1
                        'End If
                        
                        
                        'FGZ - 07/03/2006
                        If Resta_Uno Then
                            ' si es el ultimo mes deberia restarle 1 dia
                            If (MesActual = Month(Aux_Fase_Hasta(I)) And AnioActual = Year(Aux_Fase_Hasta(I))) Then
                                If Aux_Meses = 1 Then
                                    Aux_Dias = 29
                                    Aux_Meses = 0
                                Else
                                    Aux_Dias = Aux_Dias - 1
                                End If
                                
                            End If
                        End If
                        ' ---------------------------------------------------------------
                        ' Detalle
                        
                        
                        'FGZ - 04/05/2006
                        'OJO !!! cuando tiene un corte de tarea en un mes y continua en el mismo mes
                        'no coincide lo que se muestra con lo que realmente es
                        'FAF - 28-06-06 - Solucionado. Ej: sale el 10-01-06 y vuelve a entrar el 15-01-06. Da 25 dias
                        StrSql = "SELECT * FROM rep_PS62_det "
                        StrSql = StrSql & " WHERE repnro = " & Aux_RepNro
                        StrSql = StrSql & " AND anio = " & Aux_Anio
                        StrSql = StrSql & " AND bpronro = " & bpronro
                        StrSql = StrSql & " AND empresa = " & Empresa
                        StrSql = StrSql & " AND IdUser = '" & IdUser & "'"
                        StrSql = StrSql & " AND fecha = " & ConvFecha(Aux_HFecha)
                        StrSql = StrSql & " AND ternro = " & rs_Empleados!Ternro
                        StrSql = StrSql & " AND pliqanio = " & AnioActual
                        StrSql = StrSql & " AND pliqmes = " & MesActual
                        OpenRecordset StrSql, rs_Det
                        If rs_Det.EOF Then
                            'Inserto en Rep_PS62_det
                            StrSql = "INSERT INTO rep_PS62_det (repnro,anio,bpronro,empresa,iduser,fecha,hora,"
                            StrSql = StrSql & " ternro,pliqanio,pliqmes,remuneracion,sac,"
                            StrSql = StrSql & " puesto,caracter,antmeses,antdias,anthoras"
                            StrSql = StrSql & " ) VALUES ("
                            StrSql = StrSql & Aux_RepNro & ","
                            StrSql = StrSql & Aux_Anio & ","
                            StrSql = StrSql & bpronro & ","
                            StrSql = StrSql & Empresa & ","
                            StrSql = StrSql & "'" & IdUser & "',"
                            'StrSql = StrSql & ConvFecha(Fecha) & ","
                            'FGZ - 28/11/2005
                            StrSql = StrSql & ConvFecha(Aux_HFecha) & ","
                            StrSql = StrSql & "'" & hora & "',"
                            
                            StrSql = StrSql & rs_Empleados!Ternro & ","
                            StrSql = StrSql & AnioActual & ","
                            StrSql = StrSql & MesActual & ","
                            StrSql = StrSql & Aux_Acum_Remu & ","
                            StrSql = StrSql & Aux_Acum_Sac & ","
                            
                            StrSql = StrSql & "'" & Aux_Puesto & "',"
                            StrSql = StrSql & "'" & Aux_Caracter & "',"
                            StrSql = StrSql & Aux_Meses & ","
                            StrSql = StrSql & Aux_Dias & ","
                            StrSql = StrSql & Aux_Horas
                            
                            StrSql = StrSql & ")"
                            Flog.writeline "    SQL ==> " & StrSql
                            objConn.Execute StrSql, , adExecuteNoRecords
                        Else
                            '-----------------------------------------------------------------------------------------------
                            '-----------------------------------------------------------------------------------------------
                            '-----------------------------------------------------------------------------------------------
                            '17/01/2008 - Martin Ferraro - El control de sumar era correcto, pero no hacia conversion de dias
                            'a meses, ejem: si salio el 09/08/2005 y entra el 10/08/2005 muestra q trabajo 31 en vez de hacer
                            'la conversion a 1 mes
                            '-----------------------------------------------------------------------------------------------
                            '-----------------------------------------------------------------------------------------------
                            '-----------------------------------------------------------------------------------------------
'                            StrSql = "UPDATE rep_PS62_det SET "
'                            StrSql = StrSql & " remuneracion = remuneracion + " & Aux_Acum_Remu
'                            StrSql = StrSql & " ,sac = sac + " & Aux_Acum_Sac
'                            StrSql = StrSql & " ,antmeses = antmeses + " & Aux_Meses
'                            StrSql = StrSql & " ,antdias = antdias + " & Aux_Dias
'                            StrSql = StrSql & " ,anthoras = anthoras + " & Aux_Horas
'                            StrSql = StrSql & " WHERE repnro = " & Aux_RepNro
'                            StrSql = StrSql & " AND anio = " & Aux_Anio
'                            StrSql = StrSql & " AND bpronro = " & bpronro
'                            StrSql = StrSql & " AND empresa = " & Empresa
'                            StrSql = StrSql & " AND IdUser = '" & IdUser & "'"
'                            StrSql = StrSql & " AND fecha = " & ConvFecha(Aux_HFecha)
'                            StrSql = StrSql & " AND ternro = " & rs_Empleados!ternro
'                            StrSql = StrSql & " AND pliqanio = " & AnioActual
'                            StrSql = StrSql & " AND pliqmes = " & MesActual
'                            Flog.writeline "          SQL ==> " & StrSql
'                            objConn.Execute StrSql, , adExecuteNoRecords
                            
                            
                            'Acumulo los dias ya almacenados
                            Aux_Dias = Aux_Dias + rs_Det!antdias
                            
                            'Ajusto a meses segun lo trabajado considerando lo acumulado
                            If MesActual = 2 Then 'Febrero
                                If Biciesto(AnioActual) Then
                                    If Aux_Dias >= 29 Then
                                        Aux_Dias = 0
                                        Aux_Meses = 1
                                    End If
                                Else
                                    If Aux_Dias >= 28 Then
                                        Aux_Dias = 0
                                        Aux_Meses = 1
                                    End If
                                End If
                            Else
                                If MesActual = 1 Or MesActual = 3 Or MesActual = 5 Or MesActual = 7 Or MesActual = 8 Or MesActual = 10 Or MesActual = 12 Then
                                    If Aux_Dias >= 31 Then
                                        Aux_Dias = 0
                                        Aux_Meses = 1
                                    End If
                                Else
                                    If Aux_Dias >= 29 Then
                                        Aux_Dias = 0
                                        Aux_Meses = 1
                                    End If
                                End If
                            End If
                                                    
                            'Inserto el valor ajustado a meses
                            StrSql = "UPDATE rep_PS62_det SET "
                            StrSql = StrSql & " remuneracion = remuneracion + " & Aux_Acum_Remu
                            StrSql = StrSql & " ,sac = sac + " & Aux_Acum_Sac
                            StrSql = StrSql & " ,antmeses = " & Aux_Meses
                            StrSql = StrSql & " ,antdias = " & Aux_Dias
                            StrSql = StrSql & " ,anthoras = " & Aux_Horas
                            StrSql = StrSql & " WHERE repnro = " & Aux_RepNro
                            StrSql = StrSql & " AND anio = " & Aux_Anio
                            StrSql = StrSql & " AND bpronro = " & bpronro
                            StrSql = StrSql & " AND empresa = " & Empresa
                            StrSql = StrSql & " AND IdUser = '" & IdUser & "'"
                            StrSql = StrSql & " AND fecha = " & ConvFecha(Aux_HFecha)
                            StrSql = StrSql & " AND ternro = " & rs_Empleados!Ternro
                            StrSql = StrSql & " AND pliqanio = " & AnioActual
                            StrSql = StrSql & " AND pliqmes = " & MesActual
                            Flog.writeline "          SQL ==> " & StrSql
                            objConn.Execute StrSql, , adExecuteNoRecords
                            
                            'Fin correccion Martin Ferraro
                            '-----------------------------------------------------------------------------------------------
                            '-----------------------------------------------------------------------------------------------
                            '-----------------------------------------------------------------------------------------------
                        
                        End If
                    End If
                                            
                                            
                    '02/12/2005 - Fapi
                    'cambio en el metodo de calcular los años que se borran
 '                   Ultimo_Anio_Insertado = Aux_Anio
                    
                    If MesActual <> 12 Then
                        MesActual = MesActual + 1
                    Else
                        MesActual = 1
                        Aux_Anio = Aux_Anio + 1
                        AnioActual = AnioActual + 1
                    End If
                    
                    If (MesActual > Month(Aux_Fase_Hasta(I)) And AnioActual = Year(Aux_Fase_Hasta(I))) Or AnioActual > Year(Aux_Fase_Hasta(I)) Then
                        Continuar = False
                    End If
                Loop
            End If
        Next I
    
        'Actualizo el progreso del Proceso
        Progreso = Progreso + IncPorc
        TiempoAcumulado = GetTickCount
        StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & _
                 ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & _
                 "' WHERE bpronro = " & NroProcesoBatch
        objconnProgreso.Execute StrSql, , adExecuteNoRecords
        
        
        'Calculo los indices inicio y fin siguientes
        Indice_Inicio = Indice_Fin + 1
        Indice_Fin = Indice_Fin + 5
        If Indice_Fin > CantidadFases Then
            Indice_Fin = CantidadFases
        End If
        
    'siguiente grupo
    Next J
    
    rs_Empleados.MoveNext
Loop

'Reviso que queden los datos de los ultimos 10 años
'StrSql = "DELETE FROM rep_ps62_det "
'StrSql = StrSql & " WHERE empresa = " & Empresa
'StrSql = StrSql & " AND iduser = '" & IdUser & "'"
'StrSql = StrSql & " AND repnro = " & Aux_RepNro
'StrSql = StrSql & " AND anio <= " & (Ultimo_Anio_Insertado - 11) 'deja los ultimos 11 años
'objConn.Execute StrSql, , adExecuteNoRecords


'Fin de la transaccion
MyCommitTrans

If rs_Confrep.State = adStateOpen Then rs_Confrep.Close
If rs_Acu_Mes.State = adStateOpen Then rs_Acu_Mes.Close
If rs_Empleados.State = adStateOpen Then rs_Empleados.Close
If rs_Fases.State = adStateOpen Then rs_Fases.Close
If rs_Cuil.State = adStateOpen Then rs_Cuil.Close
If rs_Estructura.State = adStateOpen Then rs_Estructura.Close
If rs_Doc.State = adStateOpen Then rs_Doc.Close
If rs_Cuit.State = adStateOpen Then rs_Cuit.Close
If rs_Empresa.State = adStateOpen Then rs_Empresa.Close
If rs_Zona.State = adStateOpen Then rs_Zona.Close
If rs_Provincia.State = adStateOpen Then rs_Provincia.Close
If rs_Telefono.State = adStateOpen Then rs_Telefono.Close
If rs_Estr_Cod.State = adStateOpen Then rs_Estr_Cod.Close
If rs_HisEstructura.State = adStateOpen Then rs_HisEstructura.Close
If rs_Licencias.State = adStateOpen Then rs_Licencias.Close
If rs_rep_PS62.State = adStateOpen Then rs_rep_PS62.Close
If rs_Det.State = adStateOpen Then rs_Det.Close

Set rs_Confrep = Nothing
Set rs_Acu_Mes = Nothing
Set rs_Empleados = Nothing
Set rs_Fases = Nothing
Set rs_Cuil = Nothing
Set rs_Doc = Nothing
Set rs_Estructura = Nothing
Set rs_Doc = Nothing
Set rs_Cuit = Nothing
Set rs_Empresa = Nothing
Set rs_Zona = Nothing
Set rs_Provincia = Nothing
Set rs_Telefono = Nothing
Set rs_Estr_Cod = Nothing
Set rs_HisEstructura = Nothing
Set rs_Licencias = Nothing
Set rs_rep_PS62 = Nothing
Set rs_Det = Nothing

Exit Sub
CE:
    HuboError = True
    Flog.writeline "Error: " & Err.Description
    Flog.writeline "Ultimo sql Ejecutado: " & StrSql
    MyRollbackTrans
End Sub

Public Sub LevantarParamteros(ByVal bpronro As Long, ByVal parametros As String)
' --------------------------------------------------------------------------------------------
' Descripcion: Procedimiento para levantar los parametros pasados en batch_proceso en bprcparam
' Autor      : FGZ
' Fecha      :
' Ult. Mod   :
' Fecha      :
' --------------------------------------------------------------------------------------------
Dim pos1 As Integer
Dim pos2 As Integer
Dim aux As String

Dim Empresa As Long
Dim HFecha As Date
Dim Aux_Separador As String

Aux_Separador = "@"
' Levanto cada parametro por separado, el separador de parametros es Aux_Separador
If Not IsNull(parametros) Then
    If Len(parametros) >= 1 Then
        pos1 = 1
        pos2 = InStr(pos1, parametros, Aux_Separador) - 1
        Empresa = Mid(parametros, pos1, pos2 - pos1 + 1)
        
        pos1 = pos2 + 2
        pos2 = InStr(pos1, parametros, Aux_Separador) - 1
        HFecha = CDate(Mid(parametros, pos1, pos2 - pos1 + 1))
        
        'nombre
        pos1 = pos2 + 2
        pos2 = InStr(pos1, parametros, Aux_Separador) - 1
        aux = Mid(parametros, pos1, pos2 - pos1 + 1)
        Aux_Autoriz_Apenom = aux
        'apellido
        pos1 = pos2 + 2
        pos2 = InStr(pos1, parametros, Aux_Separador) - 1
        aux = Mid(parametros, pos1, pos2 - pos1 + 1)
        Aux_Autoriz_Apenom = aux & " " & Aux_Autoriz_Apenom
        Aux_Autoriz_Apenom = IIf(Not IsNull(Aux_Autoriz_Apenom), Aux_Autoriz_Apenom, " ")
        
        'Tipo doc
        pos1 = pos2 + 2
        pos2 = InStr(pos1, parametros, Aux_Separador) - 1
        aux = Mid(parametros, pos1, pos2 - pos1 + 1)
        Aux_Autoriz_Docu = aux
        'nro de doc
        pos1 = pos2 + 2
        pos2 = InStr(pos1, parametros, Aux_Separador) - 1
        aux = Mid(parametros, pos1, pos2 - pos1 + 1)
        Aux_Autoriz_Docu = Aux_Autoriz_Docu & "-" & aux
        Aux_Autoriz_Docu = IIf(Not IsNull(Aux_Autoriz_Docu), Aux_Autoriz_Docu, " ")
        
        pos1 = pos2 + 2
        pos2 = InStr(pos1, parametros, Aux_Separador) - 1
        Aux_Autoriz_Prov_Emis = Mid(parametros, pos1, pos2 - pos1 + 1)
        Aux_Autoriz_Prov_Emis = IIf(Not IsNull(Aux_Autoriz_Prov_Emis), Aux_Autoriz_Prov_Emis, " ")
        
        'Autoriza
        pos1 = pos2 + 2
        pos2 = InStr(pos1, parametros, Aux_Separador) - 1
        Aux_Certifi_Corresponde = Mid(parametros, pos1, pos2 - pos1 + 1)
        Aux_Certifi_Corresponde = IIf(Not IsNull(Aux_Certifi_Corresponde), Aux_Certifi_Corresponde, " ")
        
        'Tipo doc
        pos1 = pos2 + 2
        pos2 = InStr(pos1, parametros, Aux_Separador) - 1
        Aux_Certifi_Doc_Tipo = Mid(parametros, pos1, pos2 - pos1 + 1)
        Aux_Certifi_Doc_Tipo = IIf(Not IsNull(Aux_Certifi_Doc_Tipo), Aux_Certifi_Doc_Tipo, " ")
        
        'nro de doc
        pos1 = pos2 + 2
        pos2 = InStr(pos1, parametros, Aux_Separador) - 1
        Aux_Certifi_Doc_Nro = Mid(parametros, pos1, pos2 - pos1 + 1)
        Aux_Certifi_Doc_Nro = IIf(Not IsNull(Aux_Certifi_Doc_Nro), Aux_Certifi_Doc_Nro, " ")
        Aux_Certifi_Doc_Nro = Replace(Aux_Certifi_Doc_Nro, ".", "")
        
        pos1 = pos2 + 2
        pos2 = Len(parametros)
        Aux_Certifi_Expedida = Mid(parametros, pos1, pos2 - pos1 + 1)
        Aux_Certifi_Expedida = IIf(Not IsNull(Aux_Certifi_Expedida), Aux_Certifi_Expedida, " ")
        
    End If
End If

'Certificado Ansses de Servicios
Call Generar_Reporte(bpronro, Empresa, HFecha)
End Sub


'Public Function EsElUltimoEmpleado(ByVal rs As ADODB.Recordset, ByVal Anterior As Long) As Boolean
'' --------------------------------------------------------------------------------------------
'' Descripcion: Procedimiento para saber si es el ultimo empleado de la secuencia
'' Autor      : FGZ
'' Fecha      :
'' Ult. Mod   :
'' Fecha      :
'' --------------------------------------------------------------------------------------------
'
'    rs.MoveNext
'    If rs.EOF Then
'        EsElUltimoEmpleado = True
'    Else
'        If rs!Empleado <> Anterior Then
'            EsElUltimoEmpleado = True
'        Else
'            EsElUltimoEmpleado = False
'        End If
'    End If
'    rs.MovePrevious
'End Function



Public Function Antiguedad(ByRef Dia As Integer, ByRef Mes As Integer, ByRef Anio As Integer, ByRef DiasHabiles As Integer) As Integer
' -----------------------------------------------------------------------------------
' Descripcion: Antigued.p. Calcula la antiguedad al dia de hoy de un empleado en :
'               dias hábiles(si es menor que un año) o en dias, meses y años en caso contrario.
'               Retorna 0 si no hubo error y <> 0 en caso contrario
' Autor: FGZ
' Fecha: 31/07/2003
' Ultima Modificacion:
' -----------------------------------------------------------------------------------
Dim aux1 As Long
Dim aux2 As Long
Dim aux3 As Long
Dim fecalta As Date
Dim fecbaja As Date
Dim Seguir As Date
Dim q As Long
Dim NombreCampo As String
Dim rs_Fases As New ADODB.Recordset

DiasHabiles = 0

StrSql = "SELECT * FROM fases WHERE empleado = " & buliq_empleado!Ternro
OpenRecordset StrSql, rs_Fases

If rs_Fases.EOF Then
    ' ERROR. El empleado no tiene fecha de alta en fases
    Antiguedad = 1
    Exit Function
Else
        fecalta = rs_Fases!altfec
        ' verificar si se trata de un registro completo(alta/baja) o solo de un alta
        If CBool(rs_Fases!estado) Then
            fecbaja = Date  ' solo es un alta ==> tomar el Today (Date)
        Else
            fecbaja = rs_Fases!bajfec   'se trata de un registro completo
        End If
        
        Call Dif_Fechas(fecalta, fecbaja, aux1, aux2, aux3)
        Dia = Dia + aux1
        Mes = Mes + aux2 + Int(Dia / 30)
        Anio = Anio + aux3 + Int(Mes / 12)
        Dia = Dia Mod 30
        Mes = Mes Mod 12
        
        If Anio = 0 Then
            Call DiasTrab(fecalta, fecbaja, aux1)
            DiasHabiles = DiasHabiles + aux1
        End If
        Antiguedad = 0
End If

If Anio <> 0 Then
    DiasHabiles = 0
End If

' Cierro todo y Libero
If rs_Fases.State = adStateOpen Then rs_Fases.Close
Set rs_Fases = Nothing

End Function



Public Sub DiasTrab(ByVal Desde As Date, ByVal Hasta As Date, ByRef DiasH As Long)
' ---------------------------------------------------------------------------------------------
' Descripcion: Calcula la cantidad de dias trabajados de acuerdo al turno en que se trabaja y
'              de acuerdo a los dias que figuran como feriados en la tabla de feriados.
' Autor      : FGZ
' Fecha      :
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------

Dim d1 As Integer
Dim d2 As Integer
Dim aux As Integer
Dim aux2 As Integer
Dim dxsem As Integer

Dim rs_pais As New ADODB.Recordset
Dim rs_feriados As New ADODB.Recordset

    dxsem = 5
    
    d1 = Weekday(Desde)
    d2 = Weekday(Hasta)
    
    aux = DateDiff("d", Desde, Hasta) + 1
    If aux < 7 Then
        DiasH = Minimo(aux, dxsem)
    Else
        If aux = 7 Then
            DiasH = dxsem
        Else
            aux2 = 8 - d1 + d2
            If aux2 < 7 Then
                aux2 = Minimo(aux2, dxsem)
            Else
                If aux2 = 7 Then
                    aux2 = dxsem
                End If
            End If
            
            If aux2 >= 7 Then
                aux2 = Abs(aux2 - 7) + Int(aux2 / 7) * dxsem
            Else
                aux2 = aux2 + Int((aux2 - aux2) / 7) * dxsem
            End If
        End If
    End If
    
    aux = 0
    
    StrSql = "SELECT * FROM pais INNER JOIN tercero ON tercero.paisnro = pais.paisnro WHERE tercero.ternro = " & buliq_empleado!Ternro
    OpenRecordset StrSql, rs_pais
    
    If Not rs_pais.EOF Then
        ' Resto los Feriados Nacionales
        StrSql = "SELECT * FROM feriado WHERE tipferinro = 2 " & _
                 " AND fericodext = " & rs_pais!paisnro & _
                 " AND ferifecha >= " & ConvFecha(Desde) & _
                 " AND ferifecha < " & ConvFecha(Hasta)
        OpenRecordset StrSql, rs_feriados
        
        Do While Not rs_feriados.EOF
            If Weekday(rs_feriados!ferifecha) > 1 Then
                DiasH = DiasH - 1
            End If
            
            ' Siguiente Feriado
            rs_feriados.MoveNext
        Loop
    End If


    ' Resto los feriados por Convenio
    StrSql = "SELECT * FROM empleado INNER JOIN his_estructura ON empleado.ternro = his_estructura.ternro " & _
             " INNER JOIN fer_estr ON fer_estr.tenro = his_estructura.tenro " & _
             " INNER JOIN feriado ON fer_estr.ferinro = feriado.ferinro " & _
             " WHERE empleado.ternro = " & buliq_empleado!Ternro & _
             " AND feriado.tipferinro = 2" & _
             " AND feriado.ferifecha >= " & ConvFecha(Desde) & _
             " AND feriado.ferifecha < " & ConvFecha(Hasta)
    OpenRecordset StrSql, rs_feriados
    
    Do While Not rs_feriados.EOF
        If Weekday(rs_feriados!ferifecha) > 1 Then
            DiasH = DiasH - 1
        End If
        
        ' Siguiente Feriado
        rs_feriados.MoveNext
    Loop
    
    
    ' cierro todo y libero
    If rs_pais.State = adStateOpen Then rs_pais.Close
    If rs_feriados.State = adStateOpen Then rs_feriados.Close
        
    Set rs_feriados = Nothing
    Set rs_pais = Nothing

End Sub

Sub DiferenciaFase(ByVal fechaalta As Date, ByVal fechabaja As Date, ByRef diaF As Long, ByRef mesF As Long, ByRef anioF As Long)
    Dim dia1, mes1, anio1, dia2, mes2, anio2, diames As Integer
    
    dia1 = Day(fechaalta)
    mes1 = Month(fechaalta)
    anio1 = Year(fechaalta)
    dia2 = Day(fechabaja)
    mes2 = Month(fechabaja)
    anio2 = Year(fechabaja)
    
    anioF = anio2 - anio1
    
    If mes2 > mes1 Then
        mesF = mes2 - mes1 - 1
    End If
    
    If mes2 < mes1 Then
        mesF = 12 + mes2 - mes1 - 1
        anioF = anioF - 1
    End If
    
    If mes2 = mes1 Then
        mesF = -1
    End If
    
    If (mes1 = 2) And ((anio1 Mod 4) = 0) Then
        diames = 29
    End If
    
    If (mes1 = 2) And ((anio1 Mod 4) > 0) Then
        diames = 28
    End If
    
    If (mes1 = 4) Or (mes1 = 6) Or (mes1 = 9) Or (mes1 = 11) Then
        diames = 30
    End If
                
    If (mes1 = 1) Or (mes1 = 3) Or (mes1 = 5) Or (mes1 = 7) Or (mes1 = 8) Or (mes1 = 10) Or (mes1 = 12) Then
        diames = 31
    End If
    
    If (mes2 = 2) And ((anio2 Mod 4) = 0) And (dia2 = 29) Then
        dia2 = 0
        mesF = mesF + 1
    End If
    
    If (mes2 = 2) And ((anio2 Mod 4) > 0) And (dia2 = 28) Then
        dia2 = 0
        mesF = mesF + 1
    End If
    
    If ((mes2 = 4) Or (mes2 = 6) Or (mes2 = 9) Or (mes2 = 11)) And (dia2 = 30) Then
        dia2 = 0
        mesF = mesF + 1
    End If
                
    If ((mes2 = 1) Or (mes2 = 3) Or (mes2 = 5) Or (mes2 = 7) Or (mes2 = 8) Or (mes2 = 10) Or (mes2 = 12)) And (dia2 = 31) Then
        dia2 = 0
        mesF = mesF + 1
    End If
    
    diaF = dia2 + diames - dia1 + 1
    
    If (diaF >= 29) And (diames = 29) Then
        mesF = mesF + 1
        diaF = diaF - 29
    End If
    If (diaF >= 28) And (diames = 28) Then
        mesF = mesF + 1
        diaF = diaF - 28
    End If
    If (diames = 30) And (diaF >= 30) Then
        mesF = mesF + 1
        diaF = diaF - 30
    End If
    If (diames = 31) And (diaF >= 31) Then
        mesF = mesF + 1
        diaF = diaF - 31
    End If
    If mesF >= 12 Then
        mesF = mesF - 12
        anioF = anioF + 1
    End If
    If mesF = -1 Then
        mesF = 11
        anioF = anioF - 1
    End If
    
    'dia = diaF
    'mes = mesF
    'anio = anioF
    
End Sub
