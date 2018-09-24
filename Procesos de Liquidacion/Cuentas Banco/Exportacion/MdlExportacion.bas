Attribute VB_Name = "MdlExportacion"
Option Explicit

'Global Const Version = 1
'Global Const FechaVersion = "19/08/2009"   'Encriptacion de string connection
'Global Const UltimaModificacion = "Manuel Lopez"
'Global Const UltimaModificacion1 = "Encriptacion de string connection"
'Global Const Version = 1.2
'Global Const FechaVersion = "21/10/2011"   'Encriptacion de string connection
'Global Const UltimaModificacion = "Matias Dallegro"
'Global Const UltimaModificacion1 = ""

Global Const Version = 1.3
Global Const FechaVersion = "21/04/2013"   'Se cambio el query y se puso en domicilio cabdom.tidonro =3
Global Const UltimaModificacion = "Dimatz Rafael"
Global Const UltimaModificacion1 = ""
'-----------------------------------------------------------------------------

Private Type TR_Datos_Bancarios
    Proceso As String               'String   long 2  - Valor Fijo "AH"
    Servicio As String              'String   long 4  -
    Sucursal As String              'Numerico long 4  - Valor Fijo 0002
    Legajo As String                'Numerico long 20 -
    Moneda As String                'String   long 1  - Valor Fijo "P"
    Titularidad As String           'String   long 2  - Valor Fijo "SF"
    LimiteTarjeta As String         'Numerico long 2  -
End Type

Private Type TR_Datos_Personales
    Tipo_Persona As String          'String   long 1  - Valor Fijo "F"
    Apellido As String              'String   long 40 -
    nombre As String                'String   long 40 -
    Doc_Tipo As String              'String   long 2  -
    Doc_Nro As String               'Numerico long 8  -
    IVA As String                   'String   long 3  -
    ClaveTributaria_Tipo As String  'Numerico long 1  -
    ClaveTributaria_Nro As String   'Numerico long 11 -
    Fecha_Nacimiento As String      'String   long 8  - Formato "AAAAMMDD"
    Nacionalidad As String          'String   long 2  -
    Sexo As String                  'String   long 1  - Formato "F"emenino / "M"asculino
    Estado_Civil As String          'String   long 3  -
End Type

Private Type TR_Domicilio_Particular
    Calle As String                 'String   long 30 -
    Numero As String                'String   long 6  - si no existe usar "S/N"
    Piso As String                  'String   long 3  - puede ser blanco
    Depto As String                 'String   long 4  - puede ser blanco
    Codigo_Postal As String         'Numerico long 5  -
    Localidad As String             'String   long 20 -
    Provincia_Codigo As String      'String   long 2  - Segun tabla
    Telefono As String              'String   long 15 - puede ser blanco
    Fax As String                   'String   long 15 - puede ser blanco
End Type

Private Type TR_Domicilio_Laboral
    Calle As String                 'String   long 30 -
    Numero As String                'String   long 6  - si no existe usar "S/N"
    Piso As String                  'String   long 3  - puede ser blanco
    Depto As String                 'String   long 4  - puede ser blanco
    Codigo_Postal As String         'Numerico long 5  -
    Localidad As String             'String   long 20 -
    Provincia_Codigo As String      'String   long 2  - Segun tabla
    Telefono As String              'String   long 15 - puede ser blanco
    Fax As String                   'String   long 15 - puede ser blanco
End Type

Private Type TR_Datos_Varios
    Convenio_Lecop As String        'String   long 4  -
    Filler As String                'String   long 1  -
    Cliente_Ya_Existente As String  'String   long 1  -
End Type

'Private Type TR_Empleados_Suc
'    Legajo As String
'    Tercero As Long
'    Suc As String
'End Type

Global IdUser As String
Global Fecha As Date
Global hora As String

'Adrián - Declaración de dos nuevos registros.
Global rs_Empresa As New ADODB.Recordset
Global rs_tipocod As New ADODB.Recordset

Global Fecha_Inicio_periodo As Date
Global Fecha_Fin_Periodo As Date
Global StrSql2 As String

'Registros generales
Global Datos_Bancarios As TR_Datos_Bancarios
Global Datos_Personales As TR_Datos_Personales
Global Domicilio_Particular As TR_Domicilio_Particular
Global Domicilio_Laboral As TR_Domicilio_Laboral
Global Datos_varios As TR_Datos_Varios
'Global Arr_Empleados() As TR_Empleados_Suc



Public Sub Main()
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento inicial del Proceso.
' Autor      : FGZ
' Fecha      : 07/09/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim strCmdLine
Dim Nombre_Arch As String
Dim HuboError As Boolean
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

    'Abro la conexion
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
    
    Nombre_Arch = PathFLog & "Exp_Cuentas_Bco" & "-" & NroProcesoBatch & ".log"
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Nombre_Arch, True)
    
    ' Obtengo el Process ID
    ' Obtengo el Process ID
    PID = GetCurrentProcessId
    Flog.writeline "-----------------------------------------------------------------"
    Flog.writeline "Version = " & Version
    Flog.writeline "Fecha        = " & FechaVersion
    Flog.writeline "Modificacion = " & UltimaModificacion
    Flog.writeline "-----------------------------------------------------------------"
    Flog.writeline
    Flog.writeline "PID = " & PID
    
    'Cambio el estado del proceso a Procesando y el PID
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcprogreso = 0, bprcestado = 'Procesando', bprcpid = " & PID & " WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE btprcnro = 59 AND bpronro =" & NroProcesoBatch
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
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 0) & "=================================================="
    Flog.writeline Espacios(Tabulador * 0) & "Tiempo del proceso (milisegundos): " & (TiempoFinalProceso - TiempoInicialProceso)
    Flog.writeline Espacios(Tabulador * 0) & "=================================================="
    
    If Not HuboError Then
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado' WHERE bpronro = " & NroProcesoBatch
    Else
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Error' WHERE bpronro = " & NroProcesoBatch
    End If
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    Flog.Close
    objconnProgreso.Close
    objConn.Close
    
End Sub

'Public Sub Generacion(ByVal FiltroEmpleado As String, ByVal bpronro As Long, ByVal FechaDesde As Date, ByVal FechaHasta As Date, ByVal Empresa As Long, ByVal Agrupado As Boolean, _
'    ByVal AgrupaTE1 As Boolean, ByVal Tenro1 As Long, Estrnro1 As Long, _
'    ByVal AgrupaTE2 As Boolean, ByVal Tenro2 As Long, Estrnro2 As Long, _
'    ByVal AgrupaTE3 As Boolean, ByVal Tenro3 As Long, Estrnro3 As Long)
Public Sub Generacion(ByVal bpronro As Long, ByVal FechaDesde As Date, ByVal FechaHasta As Date, ByVal Empresa As Long)
' --------------------------------------------------------------------------------------------
' Descripcion: Procedimiento de generacion del archivo de Ctas nuevas para el banco
' Autor      : FGZ
' Fecha      : 01/10/2004
' Ult. Mod   :
' Fecha      :
' --------------------------------------------------------------------------------------------
Dim strLinea As String
Dim Aux_Linea As String

Dim Nro_Reporte As Integer
Dim Conf_Ok As Boolean
Dim OK As Boolean

Dim Estructura1 As Long
Dim Estructura2 As Long

Dim rs_Confrep As New ADODB.Recordset
Dim rs_Concepto As New ADODB.Recordset
Dim rs_Doc As New ADODB.Recordset
Dim rs_Periodo As New ADODB.Recordset
Dim rs_Procesos As New ADODB.Recordset
Dim rs_Estructura As New ADODB.Recordset
Dim rs_Modelo As New ADODB.Recordset
Dim rs_IVA As New ADODB.Recordset
Dim rs_Cuil As New ADODB.Recordset
Dim rs_Tercero As New ADODB.Recordset
Dim rs_Pais As New ADODB.Recordset
Dim rs_EstadoCivil As New ADODB.Recordset
Dim rs_DetDom As New ADODB.Recordset
Dim rs_Telefono As New ADODB.Recordset
Dim rs_Provincia As New ADODB.Recordset
Dim rs_Localidad As New ADODB.Recordset
Dim rs_CtaBancaria As New ADODB.Recordset
Dim rs_Parametro As New ADODB.Recordset
Dim rs_Docu As New ADODB.Recordset

Const ForReading = 1
Const TristateFalse = 0
Dim fExport
Dim fAuxiliar
Dim directorio As String
Dim Archivo As String
Dim Intentos As Integer
Dim carpeta

Dim Aux_str As String
Dim TipoCodEmpresa As String
Dim TipoCodServicio As String
Dim TipoCodLecop As String
Dim Empresa_estr As Long

Dim Nro_Concepto As String
Dim ConcNro As Long

Dim Nro_Parametro As String
Dim Parametro As Long

Dim Limite_Tarjeta As Single

Dim Aux_Linea_Datos_Bancarios As String
Dim Aux_Linea_Datos_Personales As String
Dim Aux_Linea_Domicilio_Particular As String
Dim Aux_Linea_Domicilio_laboral As String
Dim Aux_Linea_Datos_Varios As String
Dim aux_Relleno As String
Dim Aux_Tipo_Doc As Long

Dim Columna1 As Boolean
Dim Columna2 As Boolean
Dim Columna4 As Boolean
Dim Fecha As Date

'Archivo de exportacion
StrSql = "SELECT sis_dirsalidas FROM sistema WHERE sisnro = 1 "
OpenRecordset StrSql, rs
If Not rs.EOF Then
    directorio = Trim(rs!sis_dirsalidas)
End If

StrSql = "SELECT * FROM modelo WHERE modnro = 231"
OpenRecordset StrSql, rs_Modelo
If Not rs_Modelo.EOF Then
    If Not IsNull(rs_Modelo!modarchdefault) Then
        directorio = directorio & Trim(rs_Modelo!modarchdefault)
    Else
        Flog.writeline Espacios(Tabulador * 1) & "El modelo no tiene configurada la carpeta desteino. El archivo será generado en el directorio default"
    End If
Else
    Flog.writeline Espacios(Tabulador * 1) & "No se encontró el modelo. El archivo será generado en el directorio default"
End If



Archivo = directorio & "\CtasNuevas-" & Format(Date, "dd-mm-yyyy") & ".txt"
Set fs = CreateObject("Scripting.FileSystemObject")
'Activo el manejador de errores
On Error Resume Next
Set fExport = fs.CreateTextFile(Archivo, True)
If Err.Number <> 0 Then
    Flog.writeline Espacios(Tabulador * 1) & "La carpeta Destino no existe. Se creará."
    Set carpeta = fs.CreateFolder(directorio)
    Set fExport = fs.CreateTextFile(Archivo, True)
End If
'desactivo el manejador de errores
On Error GoTo 0
Flog.writeline Espacios(Tabulador * 10) & " Archivo creado en :" & Archivo
'Configuracion del Reporte
Nro_Reporte = 112
'Columna 1 - Concepto para buscar la Novedad que tiene el limite de Tarjeta
Columna1 = False
Columna2 = False
Columna4 = False
Conf_Ok = False
StrSql = "SELECT * FROM confrep WHERE repnro = " & Nro_Reporte

Flog.writeline Espacios(Tabulador * 2) & " Busco Configuracion de Reporte " & Nro_Reporte
OpenRecordset StrSql, rs_Confrep
If rs_Confrep.EOF Then
    Flog.writeline Espacios(Tabulador * 2) & "No se encontró la configuración del Reporte nº " & Nro_Reporte
    Exit Sub
Else
    Do While Not rs_Confrep.EOF
        Select Case rs_Confrep!confnrocol
        Case 1:
            Nro_Concepto = rs_Confrep!confval
            StrSql = "SELECT * FROM concepto WHERE conccod = " & Nro_Concepto
            OpenRecordset StrSql, rs_Concepto
            If rs_Concepto.EOF Then
                Flog.writeline Espacios(Tabulador * 2) & "Columna 1. El concepto no existe"
            Else
                Columna1 = True
                ConcNro = rs_Concepto!ConcNro
            End If
        Case 2:
            Nro_Parametro = rs_Confrep!confval
            StrSql = "SELECT * FROM tipopar WHERE tpanro = " & Nro_Parametro
            OpenRecordset StrSql, rs_Parametro
            If rs_Parametro.EOF Then
                Flog.writeline Espacios(Tabulador * 2) & "Columna 2. El parametro no existe " & Nro_Parametro
            Else
                Columna2 = True
                Parametro = Nro_Parametro
            End If
        Case 4:
            Aux_Tipo_Doc = rs_Confrep!confval
            StrSql = "SELECT * FROM tipodocu WHERE tidnro = " & Aux_Tipo_Doc
            OpenRecordset StrSql, rs_Docu
            If rs_Docu.EOF Then
                Flog.writeline Espacios(Tabulador * 2) & "Columna 4. El tipo de Doc no existe " & Aux_Tipo_Doc
                Aux_Tipo_Doc = 0
            Else
                Columna4 = True
                Aux_Tipo_Doc = rs_Docu!tidnro
            End If
        Case Else
        End Select
        rs_Confrep.MoveNext
    Loop
End If

Conf_Ok = Columna1 And Columna2
If Not Conf_Ok Then
    Flog.writeline Espacios(Tabulador * 2) & "Los parametros Obligatorios no estan correctamente configurados"
    Exit Sub
End If

Call EstablecerFirmas

Empresa_estr = Empresa

' Comienzo la transaccion
MyBeginTrans

'Busco los procesos a evaluar
StrSql = "SELECT distinct(empleado.ternro) as nro, empleado.*, batch_empleado.beparam "
'If AgrupaTE1 Then
'    StrSql = StrSql & ", te1.tenro tenro1, te1.estrnro estrnro1"
'End If
'If AgrupaTE2 Then
'    StrSql = StrSql & ", te2.tenro tenro2, te2.estrnro estrnro2"
'End If
'If AgrupaTE3 Then
'    StrSql = StrSql & ", te3.tenro tenro3, te3.estrnro estrnro3"
'End If
StrSql = StrSql & "  FROM  empleado "
StrSql = StrSql & " INNER JOIN batch_empleado  ON empleado.ternro = batch_empleado.ternro "
StrSql = StrSql & " INNER JOIN fases  ON fases.empleado = empleado.ternro "
StrSql = StrSql & " INNER JOIN his_estructura  ON his_estructura.ternro = empleado.ternro and his_estructura.tenro = 10 "
StrSql = StrSql & " INNER JOIN empresa ON empresa.estrnro = his_estructura.estrnro AND empresa.estrnro =" & Empresa
'If AgrupaTE1 Then
'    StrSql = StrSql & " INNER JOIN his_estructura TE1 ON te1.ternro = empleado.ternro "
'End If
'If AgrupaTE2 Then
'    StrSql = StrSql & " INNER JOIN his_estructura TE2 ON te2.ternro = empleado.ternro "
'End If
'If AgrupaTE3 Then
'    StrSql = StrSql & " INNER JOIN his_estructura TE3 ON te3.ternro = empleado.ternro "
'End If

StrSql = StrSql & " WHERE batch_empleado.bpronro =" & NroProcesoBatch
StrSql = StrSql & " AND fases.estado = -1 AND fases.real = -1 "
StrSql = StrSql & " AND ((fases.altfec <= " & ConvFecha(FechaHasta) & ")"
'StrSql = StrSql & " AND fases.altfec >= " & ConvFecha(FechaDesde)
StrSql = StrSql & " OR (fases.altfec >= " & ConvFecha(FechaDesde) & "))"
StrSql = StrSql & " AND empresa.estrnro =" & Empresa
StrSql = StrSql & " AND (his_estructura.htetdesde <= " & ConvFecha(FechaHasta) & ") AND "
StrSql = StrSql & " ((" & ConvFecha(FechaHasta) & " <= his_estructura.htethasta) or (his_estructura.htethasta is null))"
'If AgrupaTE1 Then
'    StrSql = StrSql & " AND  te1.tenro = " & Tenro1 & " AND "
'    If Estrnro1 <> 0 Then
'        StrSql = StrSql & " te1.estrnro = " & Estrnro1 & " AND "
'    End If
'    StrSql = StrSql & " (te1.htetdesde <= " & ConvFecha(FechaHasta) & ") AND "
'    StrSql = StrSql & " ((" & ConvFecha(FechaHasta) & " <= te1.htethasta) or (te1.htethasta is null)) "
'End If
'If AgrupaTE2 Then
'    StrSql = StrSql & " AND te2.tenro = " & Tenro2 & " AND "
'    If Estrnro2 <> 0 Then
'        StrSql = StrSql & " te2.estrnro = " & Estrnro2 & " AND "
'    End If
'    StrSql = StrSql & " (te2.htetdesde <= " & ConvFecha(FechaHasta) & ") AND "
'    StrSql = StrSql & " ((" & ConvFecha(FechaHasta) & " <= te2.htethasta) or (te2.htethasta is null))  "
'End If
'If AgrupaTE3 Then
'    StrSql = StrSql & " AND te3.tenro = " & Tenro3 & " AND "
'    If Estrnro3 <> 0 Then
'        StrSql = StrSql & " te3.estrnro = " & Estrnro3 & " AND "
'    End If
'    StrSql = StrSql & " (te3.htetdesde <= " & ConvFecha(FechaHasta) & ") AND "
'    StrSql = StrSql & " ((" & ConvFecha(FechaHasta) & " <= te3.htethasta) or (te3.htethasta is null))"
'End If
StrSql = StrSql & " ORDER BY empleado.ternro"
OpenRecordset StrSql, rs_Procesos

Flog.Write Espacios(Tabulador * 2) & "SQL EMPLEADOS:" & StrSql

'seteo de las variables de progreso
Progreso = 0
CConceptosAProc = rs_Procesos.RecordCount
If CConceptosAProc = 0 Then
    CConceptosAProc = 1
    Flog.writeline Espacios(Tabulador * 1) & " Ningun empleado pasa las validaciones"
End If
IncPorc = (100 / CConceptosAProc)

'Busco algunos datos globales
'Tipo de Codigo para Servicios
StrSql = "SELECT nrocod"
StrSql = StrSql & " FROM estr_cod"
StrSql = StrSql & " INNER JOIN tipocod ON tipocod.tcodnro = estr_cod.tcodnro"
StrSql = StrSql & " WHERE (tipocod.tcodnro = 11)"
StrSql = StrSql & " AND estrnro = " & Empresa_estr
If rs_tipocod.State = adStateOpen Then rs_tipocod.Close
OpenRecordset StrSql, rs_tipocod
If rs_tipocod.EOF Then
    Flog.writeline Espacios(Tabulador * 2) & " No existe número de Servicio para esta Empresa"
    TipoCodServicio = String(4, " ")
Else
    If Len(rs_tipocod!nrocod) < 4 Then
        TipoCodServicio = rs_tipocod!nrocod & String(4 - Len(rs_tipocod!nrocod), " ")
    Else
        TipoCodServicio = Left(rs_tipocod!nrocod, 4)
    End If
End If
'Tipo de Codigo para Lecop
'StrSql = "SELECT nrocod"
'StrSql = StrSql & " FROM estr_cod"
'StrSql = StrSql & " INNER JOIN tipocod ON tipocod.tcodnro = estr_cod.tcodnro"
'StrSql = StrSql & " WHERE (tipocod.tcodnro = 34)"
'StrSql = StrSql & " AND estrnro = " & Empresa_estr
'If rs_tipocod.State = adStateOpen Then rs_tipocod.Close
'OpenRecordset StrSql, rs_tipocod
'If rs_tipocod.EOF Then
    'Flog.writeline Espacios(Tabulador * 2) & " No existe número de Lecop para esta Empresa"
    TipoCodLecop = String(4, " ")
'Else
 '   If Len(rs_tipocod!nrocod) < 4 Then
       ' TipoCodLecop = rs_tipocod!nrocod & String(4 - Len(rs_tipocod!nrocod), " ")
  '  Else
    '    TipoCodLecop = Left(rs_tipocod!nrocod, 4)
   ' End If
'End If

'seteo los valores que son fijos en los registros generales
'Datos_Bancarios
Datos_Bancarios.Proceso = "AH"
Datos_Bancarios.Servicio = TipoCodServicio
Datos_Bancarios.Sucursal = "0002"
Datos_Bancarios.Legajo = String(20, "0")
Datos_Bancarios.Moneda = "P"
Datos_Bancarios.Titularidad = "SF"
Datos_Bancarios.LimiteTarjeta = "01"

'Datos_Personales
Datos_Personales.Tipo_Persona = "F"
Datos_Personales.ClaveTributaria_Tipo = "2"

'Domicilio Particular
'Domicilio_Laboral
'Datos_varios
Datos_varios.Convenio_Lecop = TipoCodLecop
Datos_varios.Filler = Space(1)
If rs_Procesos.EOF Then
    Flog.writeline Espacios(Tabulador * 2) & "No hay ningun empleado para procesar"
End If

Do While Not rs_Procesos.EOF
        Cantidad_Warnings = 0
     Flog.writeline Espacios(Tabulador * 2) & " Busco Cuenta Bancaria"
        'Reviso que el empleado no tenga ninguna Cta Bancaria Activa
        StrSql = "SELECT * FROM ctabancaria"
        StrSql = StrSql & " INNER JOIN formapago ON ctabancaria.fpagnro = formapago.fpagnro AND formapago.fpagbanc = -1"
        StrSql = StrSql & " WHERE ctabancaria.ternro =" & rs_Procesos!Ternro
        StrSql = StrSql & " AND ctabestado = -1 "
        If rs_CtaBancaria.State = adStateOpen Then rs_CtaBancaria.Close
        OpenRecordset StrSql, rs_CtaBancaria
        If rs_CtaBancaria.EOF Then
            Flog.writeline Espacios(Tabulador * 1) & "-------------------------------------"
            Flog.writeline Espacios(Tabulador * 1) & "Exportando datos del empleado " & rs_Procesos!empleg
            Flog.writeline
             ' ------------------------
             'Datos Bancarios
             ' ------------------------
             'Proceso - Fijo AH
             'Servicio - Fijo para todo el proceso
             'Sucursal - Fijo para todo el proceso
             ' 21/10/2004 FGZ
             ' ya no mas, ahora se saca del parametro batch_empleado.beparam cargado cuando se genero el proceso
             Datos_Bancarios.Sucursal = Format_StrNro(rs_Procesos!beparam, 4, True, "0")
             
             'Codigo de Agente
             If Len(rs_Procesos!empleg) < 20 Then
                 'completo con 0's a Izquierda
                 Datos_Bancarios.Legajo = String(20 - Len(rs_Procesos!empleg), "0") & rs_Procesos!empleg
             Else
                 Datos_Bancarios.Legajo = Left(rs_Procesos!empleg, 20)
             End If
             'Moneda - Fijo para todo el proceso
             'Titularidad - Fijo para todo el proceso
             'Limite de la Tarjeta
             OK = False
             Call Buscar_Novedad(ConcNro, Parametro, rs_Procesos!Ternro, FechaDesde, FechaHasta, OK, Limite_Tarjeta)
             If Not OK Then
                 Flog.writeline Espacios(Tabulador * 2) & "No se encontró ninguna Novedad Individual (Limite de Tarjeta) "
                 Cantidad_Warnings = Cantidad_Warnings + 1
                 Datos_Bancarios.LimiteTarjeta = "16"
             Else
                 Datos_Bancarios.LimiteTarjeta = Format(Limite_Tarjeta, "00")
                 Flog.writeline Espacios(Tabulador * 2) & " Limite de Tarjeta obtenido: " & Datos_Bancarios.LimiteTarjeta
             End If
             
             ' ------------------------
             'Datos Personales
             ' ------------------------
             'Tipo de Persona - Fijo "F"
             'Apellido
             Aux_str = rs_Procesos!terape & IIf(Not IsNull(rs_Procesos!terape2), rs_Procesos!terape2, "")
             Datos_Personales.Apellido = Format_Str(Aux_str, 40, True, " ")
             
             'Nombre
             Aux_str = rs_Procesos!ternom & IIf(Not IsNull(rs_Procesos!ternom2), rs_Procesos!ternom2, "")
             Datos_Personales.nombre = Format_Str(Aux_str, 40, True, " ")
            Flog.writeline Espacios(Tabulador * 2) & "Busco Tipo y Nº Doc"
             ' Buscar el documento (Tipo y Numero)
             StrSql = " SELECT ter_doc.tidnro, ter_doc.nrodoc, tipodocu.tidcod_bco FROM tercero " & _
                      " INNER JOIN ter_doc ON tercero.ternro = ter_doc.ternro " & _
                      " INNER JOIN tipodocu ON tipodocu.tidnro = ter_doc.tidnro " & _
                      " WHERE tercero.ternro = " & rs_Procesos!Ternro & _
                      " AND tipodocu.tidnro < 4 " & _
                      " ORDER BY ter_doc.tidnro "
             If rs_Doc.State = adStateOpen Then rs_Doc.Close
             OpenRecordset StrSql, rs_Doc
             If Not rs_Doc.EOF Then
                 If Not EsNulo(rs_Doc!tidcod_bco) Then
                     Datos_Personales.Doc_Tipo = rs_Doc!tidcod_bco
                     Flog.writeline Espacios(Tabulador * 2) & "Tipo de Doc obtenido: " & Datos_Personales.Doc_Tipo
                 Else
                     Datos_Personales.Doc_Tipo = "BA"
                     Flog.writeline Espacios(Tabulador * 2) & "Tipo de Doc Nulo. Se utilizará el valor por default: " & Datos_Personales.Doc_Tipo
                     Cantidad_Warnings = Cantidad_Warnings + 1
                 End If
                 Datos_Personales.Doc_Nro = Left(CStr(rs_Doc!nrodoc), 8)
                 Flog.writeline Espacios(Tabulador * 2) & " Nro: " & Datos_Personales.Doc_Nro
             Else
                 Datos_Personales.Doc_Tipo = "BA"
                 Datos_Personales.Doc_Nro = String(8, "0")
                 Flog.writeline Espacios(Tabulador * 2) & "Error al obtener los datos del Documento. Se utilizaran valores por default: " & Datos_Personales.Doc_Tipo & " Nro: " & Datos_Personales.Doc_Nro
                 Cantidad_Warnings = Cantidad_Warnings + 1
             End If
             Flog.writeline Espacios(Tabulador * 2) & "Busco la posicion ante el IVA"
             'Busco la posicion ante el IVA
             StrSql = " SELECT * FROM tercero " & _
                      " INNER JOIN posicion ON tercero.posnro = posicion.posnro " & _
                      " WHERE tercero.ternro = " & rs_Procesos!Ternro
             If rs_IVA.State = adStateOpen Then rs_IVA.Close
             OpenRecordset StrSql, rs_IVA
             If Not rs_IVA.EOF Then
                 If Not EsNulo(rs_IVA!poscod_bco) Then
                     Datos_Personales.IVA = Format_Str(rs_IVA!poscod_bco, 3, True, " ")
                     Flog.writeline Espacios(Tabulador * 2) & "Posicion ante el IVA obtenido: " & Datos_Personales.IVA
                 Else
                     Datos_Personales.IVA = "CFI"
                     Flog.writeline Espacios(Tabulador * 2) & "Posicion ante el IVA Nulo. Se utilizará el valor por default: " & Datos_Personales.IVA
                     Cantidad_Warnings = Cantidad_Warnings + 1
                 End If
             Else
                 Datos_Personales.IVA = "CFI"
                 Flog.writeline Espacios(Tabulador * 2) & "Error al obtener la Posicion ante el IVA. Se utilizaran valores por default: " & Datos_Personales.IVA
                 Cantidad_Warnings = Cantidad_Warnings + 1
             End If
             
             'Tipo Clave Tributaria - Fijo
             'Nro de Clave Tributaria (CUIL)
             StrSql = " SELECT cuil.nrodoc FROM tercero " & _
                      " INNER JOIN ter_doc cuil ON (tercero.ternro = cuil.ternro AND cuil.tidnro = 10) " & _
                      " WHERE tercero.ternro= " & rs_Procesos!Ternro
             If rs_Cuil.State = adStateOpen Then rs_Cuil.Close
             OpenRecordset StrSql, rs_Cuil
             If Not rs_Cuil.EOF Then
                 Datos_Personales.ClaveTributaria_Nro = Left(CStr(rs_Cuil!nrodoc), 13)
                 Datos_Personales.ClaveTributaria_Nro = Replace(CStr(Datos_Personales.ClaveTributaria_Nro), "-", "")
             Else
                 Flog.writeline Espacios(Tabulador * 2) & "Error al obtener el Nro de Clave tributaria (CUIL)."
                 Datos_Personales.ClaveTributaria_Nro = "00" & Datos_Personales.Doc_Nro & "0"
                 Cantidad_Warnings = Cantidad_Warnings + 1
             End If
                 
             'Fecha de Nacimiento
             StrSql = " SELECT * FROM tercero " & _
                      " WHERE tercero.ternro = " & rs_Procesos!Ternro
             If rs_Tercero.State = adStateOpen Then rs_Tercero.Close
             OpenRecordset StrSql, rs_Tercero
             If Not rs_Tercero.EOF Then
                 Datos_Personales.Fecha_Nacimiento = Format(rs_Tercero!terfecnac, "YYYYMMDD")
                 Flog.writeline Espacios(Tabulador * 2) & "Fecha de Nacimiento: " & Datos_Personales.Fecha_Nacimiento
             Else
                 Datos_Personales.Fecha_Nacimiento = Format(Date, "YYYYMMDD")
                 Flog.writeline Espacios(Tabulador * 2) & "Error al obtener la Fecha de Nacimiento. Se utilizaran valores por default: " & Datos_Personales.Fecha_Nacimiento
                 Cantidad_Warnings = Cantidad_Warnings + 1
             End If
             
             'Nacionalidad
             StrSql = " SELECT * FROM pais  " & _
                      " INNER JOIN tercero ON tercero.paisnro = pais.paisnro " & _
                      " WHERE tercero.ternro= " & rs_Procesos!Ternro
             If rs_Pais.State = adStateOpen Then rs_Pais.Close
             OpenRecordset StrSql, rs_Pais
             If Not rs_Pais.EOF Then
                 If Not EsNulo(rs_Pais!paiscod_bco) Then
                     Datos_Personales.Nacionalidad = Format_Str(rs_Pais!paiscod_bco, 2, True, " ")
                     Flog.writeline Espacios(Tabulador * 2) & "Nacionalidad obtenida: " & Datos_Personales.Nacionalidad
                 Else
                     Datos_Personales.Nacionalidad = "AR"
                     Flog.writeline Espacios(Tabulador * 2) & "Nacionalidad Nula. Se utilizará el valor por default: " & Datos_Personales.Nacionalidad
                     Cantidad_Warnings = Cantidad_Warnings + 1
                 End If
             Else
                 Datos_Personales.Nacionalidad = "AR"
                 Flog.writeline Espacios(Tabulador * 2) & "Error al obtener la Nacionalidad. Se utilizaran valores por default: " & Datos_Personales.Nacionalidad
                 Cantidad_Warnings = Cantidad_Warnings + 1
             End If
             
             'Sexo
             Datos_Personales.Sexo = IIf(CBool(rs_Tercero!tersex), "M", "F")
             
             'Estado Civil
             StrSql = " SELECT * FROM estcivil  " & _
                      " INNER JOIN tercero ON tercero.estcivnro = estcivil.estcivnro " & _
                      " WHERE tercero.ternro= " & rs_Procesos!Ternro
             If rs_EstadoCivil.State = adStateOpen Then rs_EstadoCivil.Close
             OpenRecordset StrSql, rs_EstadoCivil
             If Not rs_EstadoCivil.EOF Then
                 If Not EsNulo(rs_EstadoCivil!extciv_bco) Then
                     Datos_Personales.Estado_Civil = Format_Str(rs_EstadoCivil!extciv_bco, 1, True, " ")
                     Flog.writeline Espacios(Tabulador * 2) & "Estado Civil obtenido: " & Datos_Personales.Estado_Civil
                 Else
                     Datos_Personales.Estado_Civil = "S"
                     Flog.writeline Espacios(Tabulador * 2) & "Estado Civil Nulo. Se utilizará el valor por default: " & Datos_Personales.Estado_Civil
                     Cantidad_Warnings = Cantidad_Warnings + 1
                 End If
             Else
                 Datos_Personales.Estado_Civil = "S"
                 Flog.writeline Espacios(Tabulador * 2) & "Error al obtener el Estado Civil. Se utilizaran valores por default: " & Datos_Personales.Estado_Civil
                 Cantidad_Warnings = Cantidad_Warnings + 1
             End If
             
             ' ------------------------
             'Domicilio Particular
             ' ------------------------
             'Busco el Domicilio
             StrSql = " SELECT * FROM cabdom " & _
                      " INNER JOIN detdom ON detdom.domnro = cabdom.domnro " & _
                      " WHERE cabdom.ternro = " & rs_Procesos!Ternro & " AND " & _
                      " cabdom.tipnro =1"
             If rs_DetDom.State = adStateOpen Then rs_DetDom.Close
             OpenRecordset StrSql, rs_DetDom
             Flog.Write Espacios(Tabulador * 2) & "SQL DOMICILIOS " & StrSql
             If Not rs_DetDom.EOF Then
                 'Calle
                 Domicilio_Particular.Calle = IIf(Not IsNull(rs_DetDom!Calle), Format_Str(rs_DetDom!Calle, 30, True, " "), Space(30))
                 'Numero
                 Domicilio_Particular.Numero = IIf(Not IsNull(rs_DetDom!nro), Format_Str(rs_DetDom!nro, 6, True, " "), "S/N" & Space(3))
                 'Piso
                 Domicilio_Particular.Piso = IIf(Not IsNull(rs_DetDom!Piso), Format_Str(rs_DetDom!Piso, 3, True, " "), Space(3))
                 'Departamento
                 Domicilio_Particular.Depto = IIf(Not IsNull(rs_DetDom!oficdepto), Format_Str(rs_DetDom!oficdepto, 4, True, " "), Space(4))
                 'Codigo Postal
                 Domicilio_Particular.Codigo_Postal = IIf(Not EsNulo(rs_DetDom!codigopostal), Format_StrNro(rs_DetDom!codigopostal, 5, True, "0"), String(5, "0"))
             
                 Flog.writeline Espacios(Tabulador * 3) & "Domicilio Obtenido: "
                 Flog.writeline Espacios(Tabulador * 3) & "Calle: " & Domicilio_Particular.Calle
                 Flog.writeline Espacios(Tabulador * 3) & "Numero: " & Domicilio_Particular.Numero
                 Flog.writeline Espacios(Tabulador * 3) & "Piso: " & Domicilio_Particular.Piso
                 Flog.writeline Espacios(Tabulador * 3) & "Depto: " & Domicilio_Particular.Depto
                 Flog.writeline Espacios(Tabulador * 3) & "Codigo Postal: " & Domicilio_Particular.Codigo_Postal
                 
                 'Localidad
                 StrSql = " SELECT * FROM localidad WHERE localidad.locnro = " & rs_DetDom!locnro
                 If rs_Localidad.State = adStateOpen Then rs_Localidad.Close
                 OpenRecordset StrSql, rs_Localidad
                 If Not rs_Localidad.EOF Then
                     Domicilio_Particular.Localidad = Format_Str(rs_Localidad!locdesc, 20, True, " ")
                     Flog.writeline Espacios(Tabulador * 3) & "Localidad obtenida: " & Domicilio_Particular.Localidad
                 Else
                     Domicilio_Particular.Localidad = Space(20)
                     Flog.writeline Espacios(Tabulador * 3) & " No se encontro los datos de la Localidad."
                     Cantidad_Warnings = Cantidad_Warnings + 1
                 End If
                 
                 'Provincia
                 StrSql = " SELECT * FROM provincia WHERE provincia.provnro = " & rs_DetDom!provnro
                 If rs_Provincia.State = adStateOpen Then rs_Provincia.Close
                 OpenRecordset StrSql, rs_Provincia
                 If Not rs_Provincia.EOF Then
                     If Not EsNulo(rs_Provincia!provcod_bco) Then
                         Domicilio_Particular.Provincia_Codigo = Format_Str(rs_Provincia!provcod_bco, 2, True, " ")
                         Flog.writeline Espacios(Tabulador * 3) & "Provincia Obtenida: " & Domicilio_Laboral.Provincia_Codigo
                     Else
                         Domicilio_Particular.Provincia_Codigo = "BA"
                         Flog.writeline Espacios(Tabulador * 3) & "Codigo de Provincia Nulo. Se utilizará el valor por default: " & Domicilio_Laboral.Provincia_Codigo
                         Cantidad_Warnings = Cantidad_Warnings + 1
                     End If
                 Else
                     Domicilio_Particular.Provincia_Codigo = "BA"
                     Flog.writeline Espacios(Tabulador * 3) & " No se encontro los datos de la Provincia. Se utilizaran valores por default: " & Domicilio_Particular.Provincia_Codigo
                     Cantidad_Warnings = Cantidad_Warnings + 1
                 End If
                 
                 'Telefono r3 busco por tipotel 1 default
                 StrSql = " SELECT * FROM telefono "
                  'WHERE telefono.domnro = " & rs_DetDom!domnro
                 'StrSql = StrSql & " AND teldefault = -1"
                 StrSql = StrSql & " INNER JOIN tipotel tp ON tp.titelnro = telefono.tipotel "
                 StrSql = StrSql & " WHERE "
                 StrSql = StrSql & " telefono.domnro = " & rs_DetDom!domnro
                 StrSql = StrSql & " AND tp.titelnro = 1  "
                 If rs_Telefono.State = adStateOpen Then rs_Telefono.Close
                 OpenRecordset StrSql, rs_Telefono
                 Flog.Write Espacios(Tabulador * 3) & " SQL Tipo Tel 1 " & StrSql
                 If Not rs_Telefono.EOF Then
                     Domicilio_Particular.Telefono = IIf(Not IsNull(rs_Telefono!telnro), Format_Str(rs_Telefono!telnro, 15, True, " "), Space(15))
                 Else
                     Domicilio_Particular.Telefono = Space(15)
                     Flog.writeline Espacios(Tabulador * 3) & "No se encontro los datos del Telefono principal "
                     Cantidad_Warnings = Cantidad_Warnings + 1
                 End If
             
                 'Fax R3 BUSCO POR TIPOTEL 3 FAX
                 StrSql = " SELECT * FROM telefono"
                 'WHERE telefono.domnro = " & rs_DetDom!domnro
                 'StrSql = StrSql & " AND telfax = -1"
                  StrSql = StrSql & " INNER JOIN tipotel tp ON tp.titelnro = telefono.tipotel "
                 StrSql = StrSql & " WHERE "
                 StrSql = StrSql & " telefono.domnro = " & rs_DetDom!domnro
                 StrSql = StrSql & " AND tp.titelnro = 3 "
                 If rs_Telefono.State = adStateOpen Then rs_Telefono.Close
                 OpenRecordset StrSql, rs_Telefono
                 Flog.Write Espacios(Tabulador * 3) & " SQL Tipo Tel 3 " & StrSql
                 If Not rs_Telefono.EOF Then
                     Domicilio_Particular.Fax = IIf(Not IsNull(rs_Telefono!telnro), Format_Str(rs_Telefono!telnro, 15, True, " "), Space(15))
                 Else
                     Domicilio_Particular.Fax = Space(15)
                     Flog.writeline Espacios(Tabulador * 3) & "NO se encontro los datos del Fax "
                     Cantidad_Warnings = Cantidad_Warnings + 1
                 End If
             Else
                 Domicilio_Particular.Calle = Space(30)
                 Domicilio_Particular.Numero = "S/N" & Space(3)
                 Domicilio_Particular.Piso = Space(3)
                 Domicilio_Particular.Depto = Space(4)
                 Domicilio_Particular.Codigo_Postal = String(5, "0")
                 Domicilio_Particular.Localidad = Space(20)
                 Domicilio_Particular.Provincia_Codigo = "BA"
                 Domicilio_Particular.Telefono = Space(15)
                 Domicilio_Particular.Fax = Space(15)
                 Flog.writeline Espacios(Tabulador * 2) & "No se encontro los datos del Domicilio Particular. "
                 Cantidad_Warnings = Cantidad_Warnings + 1
             End If
             
             ' ------------------------
             'Domicilio Laboral
             ' ------------------------
             'Busco el Domicilio
             StrSql = " SELECT * FROM cabdom " & _
                      " INNER JOIN detdom ON detdom.domnro = cabdom.domnro " & _
                      " WHERE cabdom.ternro = " & rs_Procesos!Ternro & " AND " & _
                      " cabdom.tidonro =3"
             If rs_DetDom.State = adStateOpen Then rs_DetDom.Close
             OpenRecordset StrSql, rs_DetDom
            Flog.Write Espacios(Tabulador * 3) & " SQL Domicilio Laboral  " & StrSql
             If Not rs_DetDom.EOF Then
                 'Calle
                 Domicilio_Laboral.Calle = IIf(Not IsNull(rs_DetDom!Calle), Format_Str(rs_DetDom!Calle, 30, True, " "), Space(30))
                 'Numero
                 Domicilio_Laboral.Numero = IIf(Not IsNull(rs_DetDom!nro), Format_Str(rs_DetDom!nro, 6, True, " "), "S/N" & Space(3))
                 'Piso
                 Domicilio_Laboral.Piso = IIf(Not IsNull(rs_DetDom!Piso), Format_Str(rs_DetDom!Piso, 3, True, " "), Space(3))
                 'Departamento
                 Domicilio_Laboral.Depto = IIf(Not IsNull(rs_DetDom!oficdepto), Format_Str(rs_DetDom!oficdepto, 4, True, " "), Space(4))
                 'Codigo Postal
                 Domicilio_Laboral.Codigo_Postal = IIf(Not IsNull(rs_DetDom!cpdigopostal), Format_StrNro(rs_DetDom!codigopostal, 5, True, "0"), String(5, "0"))
             
                 Flog.writeline Espacios(Tabulador * 3) & "Domicilio Obtenido: "
                 Flog.writeline Espacios(Tabulador * 3) & "Calle: " & Domicilio_Laboral.Calle
                 Flog.writeline Espacios(Tabulador * 3) & "Numero: " & Domicilio_Laboral.Numero
                 Flog.writeline Espacios(Tabulador * 3) & "Piso: " & Domicilio_Laboral.Piso
                 Flog.writeline Espacios(Tabulador * 3) & "Depto: " & Domicilio_Laboral.Depto
                 Flog.writeline Espacios(Tabulador * 3) & "Codigo Postal: " & Domicilio_Laboral.Codigo_Postal
                 
                 'Localidad
                 StrSql = " SELECT * FROM localidad WHERE localidad.locnro = " & rs_DetDom!locnro
                 If rs_Localidad.State = adStateOpen Then rs_Localidad.Close
                 OpenRecordset StrSql, rs_Localidad
                 If Not rs_Localidad.EOF Then
                     Domicilio_Laboral.Localidad = Format_Str(rs_Localidad!locdesc, 20, True, " ")
                     Flog.writeline Espacios(Tabulador * 3) & "Localidad obtenida: " & Domicilio_Laboral.Localidad
                 Else
                     Domicilio_Laboral.Localidad = Space(20)
                     Flog.writeline Espacios(Tabulador * 3) & "No se encontro los datos de la Localidad."
                     Cantidad_Warnings = Cantidad_Warnings + 1
                 End If
                 
                 'Provincia
                 StrSql = " SELECT * FROM provincia WHERE provincia.provnro = " & rs_DetDom!provnro
                 If rs_Provincia.State = adStateOpen Then rs_Provincia.Close
                 OpenRecordset StrSql, rs_Provincia
                 If Not rs_Provincia.EOF Then
                     If Not EsNulo(rs_Provincia!provcod_bco) Then
                         Domicilio_Laboral.Provincia_Codigo = Format_Str(rs_Provincia!provcod_bco, 2, True, " ")
                         Flog.writeline Espacios(Tabulador * 3) & "Provincia Obtenida: " & Domicilio_Laboral.Provincia_Codigo
                     Else
                         Domicilio_Laboral.Provincia_Codigo = "BA"
                         Flog.writeline Espacios(Tabulador * 3) & "Codigo de Provincia Nulo. Se utilizará el valor por default: " & Domicilio_Laboral.Provincia_Codigo
                         Cantidad_Warnings = Cantidad_Warnings + 1
                     End If
                 Else
                     Domicilio_Laboral.Provincia_Codigo = "BA"
                     Flog.writeline Espacios(Tabulador * 3) & "No se encontro los datos de la Provincia. Se utilizaran valores por default: " & Domicilio_Laboral.Provincia_Codigo
                     Cantidad_Warnings = Cantidad_Warnings + 1
                 End If
                 
                 'Telefono
                 StrSql = " SELECT * FROM telefono "
                 'WHERE telefono.domnro = " & rs_DetDom!domnro
                 'StrSql = StrSql & " AND teldefault = -1"
                 StrSql = StrSql & " INNER JOIN tipotel tp ON tp.titelnro = telefono.tipotel "
                 StrSql = StrSql & " WHERE "
                 StrSql = StrSql & " telefono.domnro = " & rs_DetDom!domnro
                 StrSql = StrSql & " AND tp.titelnro = 1  "
                 If rs_Telefono.State = adStateOpen Then rs_Telefono.Close
                 OpenRecordset StrSql, rs_Telefono
                 Flog.Write Espacios(Tabulador * 3) & " SQL Telefono tipo tel 1  " & StrSql
                 If Not rs_Telefono.EOF Then
                     Domicilio_Laboral.Telefono = IIf(Not IsNull(rs_Telefono!telnro), Format_Str(rs_Telefono!telnro, 15, True, " "), Space(15))
                 Else
                     Domicilio_Laboral.Telefono = Space(15)
                     Flog.writeline Espacios(Tabulador * 3) & "No se encontro los datos del Telefono principal "
                     Cantidad_Warnings = Cantidad_Warnings + 1
                 End If
             
                 'Fax
                 StrSql = " SELECT * FROM telefono "
                 'WHERE telefono.domnro = " & rs_DetDom!domnro
                 'StrSql = StrSql & " AND telfax = -1"
                 StrSql = StrSql & " INNER JOIN tipotel tp ON tp.titelnro = telefono.tipotel "
                 StrSql = StrSql & " WHERE "
                 StrSql = StrSql & " telefono.domnro = " & rs_DetDom!domnro
                 StrSql = StrSql & " AND tp.titelnro = 3 "
                 If rs_Telefono.State = adStateOpen Then rs_Telefono.Close
                 OpenRecordset StrSql, rs_Telefono
                Flog.Write Espacios(Tabulador * 3) & " SQL Fax tipo tel 2  " & StrSql
                 If Not rs_Telefono.EOF Then
                     Domicilio_Laboral.Fax = IIf(Not IsNull(rs_Telefono!telnro), Format_Str(rs_Telefono!telnro, 15, True, " "), Space(15))
                 Else
                     Domicilio_Laboral.Fax = Space(15)
                     Flog.writeline Espacios(Tabulador * 3) & "No se encontro los datos del Fax "
                     Cantidad_Warnings = Cantidad_Warnings + 1
                 End If
             Else
                 Domicilio_Laboral.Calle = Space(30)
                 Domicilio_Laboral.Numero = "S/N" & Space(3)
                 Domicilio_Laboral.Piso = Space(3)
                 Domicilio_Laboral.Depto = Space(4)
                 Domicilio_Laboral.Codigo_Postal = String(5, "0")
                 Domicilio_Laboral.Localidad = Space(20)
                 Domicilio_Laboral.Provincia_Codigo = "BA"
                 Domicilio_Laboral.Telefono = Space(15)
                 Domicilio_Laboral.Fax = Space(15)
                 Flog.writeline Espacios(Tabulador * 2) & "No se encontro los datos del Domicilio Laboral. "
                 Cantidad_Warnings = Cantidad_Warnings + 1
             End If
             
             
             ' ------------------------
             'Datos varios
             ' ------------------------
             'Convenio en Lecop - Fijo para el proceso
             'Filler - ??????
             'Cliente Ya Existente para el Banco
             If Columna4 Then
                ' Buscar el documento (Tipo 33). Si existe ==> ya era cliente
                StrSql = " SELECT ter_doc.tidnro, ter_doc.nrodoc, tipodocu.tidcod_bco FROM tercero " & _
                         " INNER JOIN ter_doc ON tercero.ternro = ter_doc.ternro " & _
                         " INNER JOIN tipodocu ON tipodocu.tidnro = ter_doc.ternro " & _
                         " WHERE tercero.ternro = " & rs_Procesos!Ternro & _
                         " AND tipodocu.tidnro = " & Aux_Tipo_Doc & _
                         " ORDER BY ter_doc.tidnro "
                If rs_Doc.State = adStateOpen Then rs_Doc.Close
                OpenRecordset StrSql, rs_Doc
                If Not rs_Doc.EOF Then
                    Datos_varios.Cliente_Ya_Existente = "S"
                    Flog.writeline Espacios(Tabulador * 2) & "El empleado ya tiene cuenta en el Banco"
                Else
                    Flog.writeline Espacios(Tabulador * 2) & "El empleado No era cliente del Banco"
                    Datos_varios.Cliente_Ya_Existente = "N" 'Space(1)
                End If
            Else
                Flog.writeline Espacios(Tabulador * 2) & "Default. El empleado No era cliente del Banco"
                Datos_varios.Cliente_Ya_Existente = "N"
            End If
            
             ' ------------------------------------------------------------------------
             'Escribo en el archivo de texto
             
             Aux_Linea_Datos_Bancarios = Datos_Bancarios.Proceso
             Aux_Linea_Datos_Bancarios = Aux_Linea_Datos_Bancarios & Datos_Bancarios.Servicio
             Aux_Linea_Datos_Bancarios = Aux_Linea_Datos_Bancarios & Datos_Bancarios.Sucursal
             Aux_Linea_Datos_Bancarios = Aux_Linea_Datos_Bancarios & Datos_Bancarios.Legajo
             Aux_Linea_Datos_Bancarios = Aux_Linea_Datos_Bancarios & Datos_Bancarios.Moneda
             Aux_Linea_Datos_Bancarios = Aux_Linea_Datos_Bancarios & Datos_Bancarios.Titularidad
             Aux_Linea_Datos_Bancarios = Aux_Linea_Datos_Bancarios & Datos_Bancarios.LimiteTarjeta
             
             Aux_Linea_Datos_Personales = Datos_Personales.Tipo_Persona
             Aux_Linea_Datos_Personales = Aux_Linea_Datos_Personales & Datos_Personales.Apellido
             Aux_Linea_Datos_Personales = Aux_Linea_Datos_Personales & Datos_Personales.nombre
             Aux_Linea_Datos_Personales = Aux_Linea_Datos_Personales & Datos_Personales.Doc_Tipo
             Aux_Linea_Datos_Personales = Aux_Linea_Datos_Personales & Datos_Personales.Doc_Nro
             Aux_Linea_Datos_Personales = Aux_Linea_Datos_Personales & Datos_Personales.IVA
             Aux_Linea_Datos_Personales = Aux_Linea_Datos_Personales & Datos_Personales.ClaveTributaria_Tipo
             Aux_Linea_Datos_Personales = Aux_Linea_Datos_Personales & Datos_Personales.ClaveTributaria_Nro
             Aux_Linea_Datos_Personales = Aux_Linea_Datos_Personales & Datos_Personales.Fecha_Nacimiento
             Aux_Linea_Datos_Personales = Aux_Linea_Datos_Personales & Datos_Personales.Nacionalidad
             Aux_Linea_Datos_Personales = Aux_Linea_Datos_Personales & Datos_Personales.Sexo
             Aux_Linea_Datos_Personales = Aux_Linea_Datos_Personales & Datos_Personales.Estado_Civil
             
             Aux_Linea_Domicilio_Particular = Domicilio_Particular.Calle
             Aux_Linea_Domicilio_Particular = Aux_Linea_Domicilio_Particular & Domicilio_Particular.Numero
             Aux_Linea_Domicilio_Particular = Aux_Linea_Domicilio_Particular & Domicilio_Particular.Piso
             Aux_Linea_Domicilio_Particular = Aux_Linea_Domicilio_Particular & Domicilio_Particular.Depto
             Aux_Linea_Domicilio_Particular = Aux_Linea_Domicilio_Particular & Domicilio_Particular.Codigo_Postal
             Aux_Linea_Domicilio_Particular = Aux_Linea_Domicilio_Particular & Domicilio_Particular.Localidad
             Aux_Linea_Domicilio_Particular = Aux_Linea_Domicilio_Particular & Domicilio_Particular.Provincia_Codigo
             Aux_Linea_Domicilio_Particular = Aux_Linea_Domicilio_Particular & Domicilio_Particular.Telefono
             Aux_Linea_Domicilio_Particular = Aux_Linea_Domicilio_Particular & Domicilio_Particular.Fax
             
             Aux_Linea_Domicilio_laboral = Domicilio_Laboral.Calle
             Aux_Linea_Domicilio_laboral = Aux_Linea_Domicilio_laboral & Domicilio_Laboral.Numero
             Aux_Linea_Domicilio_laboral = Aux_Linea_Domicilio_laboral & Domicilio_Laboral.Piso
             Aux_Linea_Domicilio_laboral = Aux_Linea_Domicilio_laboral & Domicilio_Laboral.Depto
             Aux_Linea_Domicilio_laboral = Aux_Linea_Domicilio_laboral & Domicilio_Laboral.Codigo_Postal
             Aux_Linea_Domicilio_laboral = Aux_Linea_Domicilio_laboral & Domicilio_Laboral.Localidad
             Aux_Linea_Domicilio_laboral = Aux_Linea_Domicilio_laboral & Domicilio_Laboral.Provincia_Codigo
             Aux_Linea_Domicilio_laboral = Aux_Linea_Domicilio_laboral & Domicilio_Laboral.Telefono
             Aux_Linea_Domicilio_laboral = Aux_Linea_Domicilio_laboral & Domicilio_Laboral.Fax
             
             Aux_Linea_Datos_Varios = Datos_varios.Convenio_Lecop
             Aux_Linea_Datos_Varios = Aux_Linea_Datos_Varios & Datos_varios.Filler
             Aux_Linea_Datos_Varios = Aux_Linea_Datos_Varios & Datos_varios.Cliente_Ya_Existente
             
             aux_Relleno = Space(700 - Len(Aux_Linea_Datos_Bancarios & Aux_Linea_Datos_Personales & Aux_Linea_Domicilio_Particular & Aux_Linea_Domicilio_laboral & Aux_Linea_Datos_Varios))
             fExport.writeline Aux_Linea_Datos_Bancarios & Aux_Linea_Datos_Personales & Aux_Linea_Domicilio_Particular & Aux_Linea_Domicilio_laboral & aux_Relleno & Aux_Linea_Datos_Varios
        
            Flog.writeline Espacios(Tabulador * 1) & "-------------------------------------"
            If Cantidad_Warnings > 0 Then
                Flog.writeline Espacios(Tabulador * 1) & "Atención !!! se han detectado " & Cantidad_Warnings & " advertencias para el empleado"
                Flog.writeline Espacios(Tabulador * 1) & "Empleado exportado correctamente"
            Else
                Flog.writeline Espacios(Tabulador * 1) & "Empleado exportado correctamente"
            End If
        Else
            Flog.writeline Espacios(Tabulador * 1) & "-------------------------------------"
            Flog.writeline Espacios(Tabulador * 1) & "El empleado ya tiene una cuenta " & rs_Procesos!empleg
            Flog.writeline Espacios(Tabulador * 1) & "No se exportaran sus datos"
        End If
        
        'Actualizo el progreso del Proceso
        Progreso = Progreso + IncPorc
        TiempoAcumulado = GetTickCount
        StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & _
                 ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & _
                 "' WHERE bpronro = " & NroProcesoBatch
        objconnProgreso.Execute StrSql, , adExecuteNoRecords
                
    'Siguiente proceso
    rs_Procesos.MoveNext
Loop
'Cierro el archivo creado
fExport.Close

'Fin de la transaccion
MyCommitTrans


If rs_Confrep.State = adStateOpen Then rs_Confrep.Close
If rs_Concepto.State = adStateOpen Then rs_Concepto.Close
If rs_Doc.State = adStateOpen Then rs_Doc.Close
If rs_Procesos.State = adStateOpen Then rs_Procesos.Close
If rs_Periodo.State = adStateOpen Then rs_Periodo.Close
If rs_Estructura.State = adStateOpen Then rs_Estructura.Close
If rs_tipocod.State = adStateOpen Then rs_tipocod.Close
If rs_Empresa.State = adStateOpen Then rs_Empresa.Close
If rs_Modelo.State = adStateOpen Then rs_Modelo.Close
If rs_IVA.State = adStateOpen Then rs_IVA.Close
If rs_Cuil.State = adStateOpen Then rs_Cuil.Close
If rs_Tercero.State = adStateOpen Then rs_Tercero.Close
If rs_Pais.State = adStateOpen Then rs_Pais.Close
If rs_EstadoCivil.State = adStateOpen Then rs_EstadoCivil.Close
If rs_DetDom.State = adStateOpen Then rs_DetDom.Close
If rs_Telefono.State = adStateOpen Then rs_Telefono.Close
If rs_Provincia.State = adStateOpen Then rs_Provincia.Close
If rs_Localidad.State = adStateOpen Then rs_Localidad.Close
If rs_CtaBancaria.State = adStateOpen Then rs_CtaBancaria.Close
If rs_Parametro.State = adStateOpen Then rs_Parametro.Close
If rs_Docu.State = adStateOpen Then rs_Docu.Close


Set rs_Confrep = Nothing
Set rs_Concepto = Nothing
Set rs_Doc = Nothing
Set rs_Procesos = Nothing
Set rs_Periodo = Nothing
Set rs_Estructura = Nothing
Set rs_tipocod = Nothing
Set rs_Empresa = Nothing
Set rs_Modelo = Nothing
Set rs_IVA = Nothing
Set rs_Cuil = Nothing
Set rs_Tercero = Nothing
Set rs_Pais = Nothing
Set rs_EstadoCivil = Nothing
Set rs_DetDom = Nothing
Set rs_Telefono = Nothing
Set rs_Provincia = Nothing
Set rs_Localidad = Nothing
Set rs_CtaBancaria = Nothing
Set rs_Parametro = Nothing
Set rs_Docu = Nothing

Exit Sub
CE:
    HuboError = True
    MyRollbackTrans

If rs_Confrep.State = adStateOpen Then rs_Confrep.Close
If rs_Concepto.State = adStateOpen Then rs_Concepto.Close
If rs_Doc.State = adStateOpen Then rs_Doc.Close
If rs_Procesos.State = adStateOpen Then rs_Procesos.Close
If rs_Periodo.State = adStateOpen Then rs_Periodo.Close
If rs_Estructura.State = adStateOpen Then rs_Estructura.Close
If rs_tipocod.State = adStateOpen Then rs_tipocod.Close
If rs_Empresa.State = adStateOpen Then rs_Empresa.Close
If rs_Modelo.State = adStateOpen Then rs_Modelo.Close
If rs_IVA.State = adStateOpen Then rs_IVA.Close
If rs_Cuil.State = adStateOpen Then rs_Cuil.Close
If rs_Tercero.State = adStateOpen Then rs_Tercero.Close
If rs_Pais.State = adStateOpen Then rs_Pais.Close
If rs_EstadoCivil.State = adStateOpen Then rs_EstadoCivil.Close
If rs_DetDom.State = adStateOpen Then rs_DetDom.Close
If rs_Telefono.State = adStateOpen Then rs_Telefono.Close
If rs_Provincia.State = adStateOpen Then rs_Provincia.Close
If rs_Localidad.State = adStateOpen Then rs_Localidad.Close
If rs_CtaBancaria.State = adStateOpen Then rs_CtaBancaria.Close
If rs_Parametro.State = adStateOpen Then rs_Parametro.Close
If rs_Docu.State = adStateOpen Then rs_Docu.Close


Set rs_Confrep = Nothing
Set rs_Concepto = Nothing
Set rs_Doc = Nothing
Set rs_Procesos = Nothing
Set rs_Periodo = Nothing
Set rs_Estructura = Nothing
Set rs_tipocod = Nothing
Set rs_Empresa = Nothing
Set rs_Modelo = Nothing
Set rs_IVA = Nothing
Set rs_Cuil = Nothing
Set rs_Tercero = Nothing
Set rs_Pais = Nothing
Set rs_EstadoCivil = Nothing
Set rs_DetDom = Nothing
Set rs_Telefono = Nothing
Set rs_Provincia = Nothing
Set rs_Localidad = Nothing
Set rs_CtaBancaria = Nothing
Set rs_Parametro = Nothing
Set rs_Docu = Nothing

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
Dim Separador As String

Dim FechaDesde As Date
Dim FechaHasta As Date
Dim pliqdesde As Long
Dim pliqhasta As Long
Dim Todos_Pro As Boolean
Dim Proc_Aprob As Integer
Dim Empresa As Long
Dim FiltroEmpleados As String

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

'Orden de los parametros
'pliqdesde
'pliqhasta
'fecha desde
'fecha hasta
'Proaprob
'Lista de procesos
'Tenro1
'estrnro1
'tenro2
'estrnro2
'tenro3
'estrnro3
'Empresa empnro
'no calienta

Separador = "@"
' Levanto cada parametro por separado
If Not IsNull(parametros) Then
    If Len(parametros) >= 1 Then
        pos1 = 1
        pos2 = InStr(pos1, parametros, Separador) - 1
        FechaDesde = CDate(Mid(parametros, pos1, pos2 - pos1 + 1))
        
        pos1 = pos2 + 2
        pos2 = InStr(pos1, parametros, Separador) - 1
        FechaHasta = CDate(Mid(parametros, pos1, pos2 - pos1 + 1))
        
'        pos1 = pos2 + 2
'        pos2 = InStr(pos1, parametros, Separador) - 1
'        FiltroEmpleados = Mid(parametros, pos1, pos2 - pos1 + 1)
'        pos1 = InStr(1, FiltroEmpleados, "e") - 1
'        FiltroEmpleados = Mid(FiltroEmpleados, pos1, pos2 - pos1 + 1)
'
''        '-------------------------------------------
''            pos1 = pos2 + 2
''            pos2 = InStr(pos1, parametros, Separador) - 1
''            Cantidad_Empleados = Mid(parametros, pos1, pos2 - pos1 + 1)
''
''            ReDim Preserve Arr_Empleados(Cantidad_Empleados + 1) As TR_Empleados_Suc
''
''            pos1 = pos2 + 2
''            pos2 = InStr(pos1, parametros, Separador) - 1
''            FiltroEmpleados = Mid(parametros, pos1, pos2 - pos1 + 1)
''
''
''            For i = 1 To Cantidad_Empleados
''                pos3 = 1
''                pos4 = InStr(pos3, parametros, ",") - 1
''                aux_leg = Mid(parametros, pos3, pos4 - pos3 + 1)
''
''
''
''            Next i
''
''        '-------------------------------------------
'
'        pos1 = pos2 + 2
'        pos2 = InStr(pos1, parametros, Separador) - 1
'        Tenro1 = Mid(parametros, pos1, pos2 - pos1 + 1)
'        If Not Tenro1 = 0 Then
'            Agrupado = True
'            AgrupaTE1 = True
'        End If
'        pos1 = pos2 + 2
'        pos2 = InStr(pos1, parametros, Separador) - 1
'        Estrnro1 = Mid(parametros, pos1, pos2 - pos1 + 1)
'
'        pos1 = pos2 + 2
'        pos2 = InStr(pos1, parametros, Separador) - 1
'        Tenro2 = Mid(parametros, pos1, pos2 - pos1 + 1)
'        If Not Tenro2 = 0 Then
'            AgrupaTE2 = True
'        End If
'        pos1 = pos2 + 2
'        pos2 = InStr(pos1, parametros, Separador) - 1
'        Estrnro2 = Mid(parametros, pos1, pos2 - pos1 + 1)
'
'        pos1 = pos2 + 2
'        pos2 = InStr(pos1, parametros, Separador) - 1
'        Tenro3 = Mid(parametros, pos1, pos2 - pos1 + 1)
'        If Not Tenro3 = 0 Then
'            AgrupaTE3 = True
'        End If
'        pos1 = pos2 + 2
'        pos2 = InStr(pos1, parametros, Separador) - 1
'        Estrnro3 = Mid(parametros, pos1, pos2 - pos1 + 1)
        
        pos1 = pos2 + 2
        pos2 = Len(parametros)
        Empresa = Mid(parametros, pos1, pos2 - pos1 + 1)
        'Empresa = 1
    End If
End If
'Call Generacion(FiltroEmpleados, bpronro, pliqdesde, Todos_Pro, Proc_Aprob, Empresa, Agrupado, AgrupaTE1, Tenro1, Estrnro1, AgrupaTE2, Tenro2, Estrnro2, AgrupaTE3, Tenro3, Estrnro3)
'Call Generacion(FiltroEmpleados, bpronro, FechaDesde, FechaHasta, Empresa, Agrupado, AgrupaTE1, Tenro1, Estrnro1, AgrupaTE2, Tenro2, Estrnro2, AgrupaTE3, Tenro3, Estrnro3)
Call Generacion(bpronro, FechaDesde, FechaHasta, Empresa)
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


Public Sub Buscar_Novedad(ByVal concepto As Long, ByVal Tpanro As Long, ByVal Tercero As Long, ByVal Fecha_Inicio As Date, ByVal Fecha_Fin As Date, ByRef Encontro As Boolean, ByRef Val As Single)
' ---------------------------------------------------------------------------------------------
' Descripcion: Obtencion de novedad a Nivel Individual.
' Autor      : FGZ
' Fecha      : 07/10/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim rs_NovEmp As New ADODB.Recordset

Dim Vigencia_Activa As Boolean
Dim Firmado As Boolean
Dim rs_firmas As New ADODB.Recordset
Dim OK As Boolean

    Encontro = False
    OK = False

        StrSql = "SELECT * FROM novemp WHERE " & _
                 " concnro = " & concepto & _
                 " AND tpanro = " & Tpanro & _
                 " AND empleado = " & Tercero & _
                 " AND ((nevigencia = -1 " & _
                 " AND nedesde < " & ConvFecha(Fecha_Fin) & _
                 " AND (nehasta >= " & ConvFecha(Fecha_Inicio) & _
                 " OR nehasta is null )) " & _
                 " OR nevigencia = 0)" & _
                 " ORDER BY nevigencia, nedesde, nehasta "
        OpenRecordset StrSql, rs_NovEmp
        
        Val = 0
        Do While Not rs_NovEmp.EOF
            If FirmaActiva5 Then
                '/* Verificar si esta en el NIVEL FINAL DE FIRMA */
                StrSql = "select * from cysfirmas where cysfirfin = -1 " & _
                         " AND cysfircodext = '" & rs_NovEmp!nenro & "' and cystipnro = 5"
                OpenRecordset StrSql, rs_firmas
                If rs_firmas.EOF Then
                    Firmado = False
                    Flog.writeline Espacios(Tabulador * 2) & "NIVEL FINAL DE FIRMA No Activo "
                Else
                    Firmado = True
                End If
                If rs_firmas.State = adStateOpen Then rs_firmas.Close
            Else
                Firmado = True
            End If

        If Firmado Then
            If CBool(rs_NovEmp!nevigencia) Then
                Vigencia_Activa = True
                If Not EsNulo(rs_NovEmp!nehasta) Then
                    If (rs_NovEmp!nehasta < Fecha_Inicio) Or (Fecha_Fin < rs_NovEmp!nedesde) Then
                        Vigencia_Activa = False
                        Flog.writeline Espacios(Tabulador * 2) & "Novedad con vigencia desde " & rs_NovEmp!nedesde & " Hasta " & rs_NovEmp!nehasta & " INACTIVA con valor " & rs_NovEmp!nevalor
                    Else
                        Flog.writeline Espacios(Tabulador * 2) & "Novedad con vigencia desde " & rs_NovEmp!nedesde & " Hasta " & rs_NovEmp!nehasta & " ACTIVA con valor " & rs_NovEmp!nevalor
                    End If
                Else
                    If (Fecha_Fin < rs_NovEmp!nedesde) Then
                        Vigencia_Activa = False
                        Flog.writeline Espacios(Tabulador * 2) & "Novedad con vigencia desde " & rs_NovEmp!nedesde & " Hasta indeterminado INACTIVA con valor " & rs_NovEmp!nevalor
                    Else
                        Flog.writeline Espacios(Tabulador * 2) & "Novedad con vigencia desde " & rs_NovEmp!nedesde & " Hasta indeterminado ACTIVA con valor " & rs_NovEmp!nevalor
                    End If
                End If
            Else
                Flog.writeline Espacios(Tabulador * 2) & "Novedad sin vigencia con valor " & rs_NovEmp!nevalor
            End If
            
            If Vigencia_Activa Or Not CBool(rs_NovEmp!nevigencia) Then
                Val = Val + rs_NovEmp!nevalor
            End If
            
            OK = True
            Encontro = True
        End If 'If Firmado Then
        
        rs_NovEmp.MoveNext
    Loop
End Sub


Public Sub EstablecerFirmas()
Dim rs_cystipo As New ADODB.Recordset

    
    FirmaActiva5 = False
    FirmaActiva15 = False
    FirmaActiva19 = False
    FirmaActiva20 = False
    
    StrSql = "select * from cystipo where (cystipnro = 5 or cystipnro = 15 OR cystipnro = 19 or cystipnro = 20) AND cystipact = -1"
    OpenRecordset StrSql, rs_cystipo
    
    Do While Not rs_cystipo.EOF
    Select Case rs_cystipo!cystipnro
    Case 5:
        FirmaActiva5 = True
    Case 15:
        FirmaActiva15 = True
    Case 19:
        FirmaActiva19 = True
    Case 20:
        FirmaActiva20 = True
    Case Else
    End Select
        
        rs_cystipo.MoveNext
    Loop
    
If rs_cystipo.State = adStateOpen Then rs_cystipo.Close
Set rs_cystipo = Nothing

End Sub

