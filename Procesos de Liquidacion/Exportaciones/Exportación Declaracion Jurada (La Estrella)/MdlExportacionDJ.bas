Attribute VB_Name = "MdlExportacion"
Option Explicit

'Global Const Version = "1.01"
'Global Const FechaModificacion = "16/12/2005"
'Global Const UltimaModificacion = "" '"No funcionaban los niveles de agrupacion."

Global Const Version = "1.02"
Global Const FechaModificacion = "31/07/2009"
Global Const UltimaModificacion = "" 'Encriptacion de string connection

Private Type TipoReg1
    Tipo_Reg As String              'Numerico long 1  - Valor Fijo 1
    Nro_ID As String                'Numerico long 15 -
    Total_Aportes As Single         'Numerico long 16 - 14 enteros y 2 decimales
    Salario_MesAno As String        'Numerico long 4  - MMAA
    Total_Pag As String             'Numerico long 4  - Valor Fijo 1
    Codigo_Declaracion As String    'Numerico long 1  - Valor Fijo 1 - DJA Original
End Type
Private Type TipoReg2
    Tipo_Reg As String              'Numerico long 1  - Valor Fijo 2
    Nro_ID As String                'Numerico long 15 -
    Nro_Pag As String               'Numerico long 4  - Valor Fijo 1
    Total_Aportes As Single         'Numerico long 16 - 14 enteros y 2 decimales
    Espacios As String              'Numerico long 5  - en blanco
End Type
Private Type TipoReg3
    Tipo_Reg As String              'Numerico long 1  - Valor Fijo 3
    Nro_ID As String                'Numerico long 15 -
    Tipo_Doc As String              'Numerico long 1  - 1 - DNI / LC / LE y 4 - CI
    Nro_Doc As String               'Numerico long 8  -
    Importe As Single               'Numerico long 15 - 13 enteros y 2 decimales
    Espacios As String              'string   long 1  - en blanco
End Type

Global IdUser As String
Global Fecha As Date
Global Hora As String

'Adrián - Declaración de dos nuevos registros.
Global rs_Empresa As New ADODB.Recordset
Global rs_tipocod As New ADODB.Recordset

Global Fecha_Inicio_periodo As Date
Global Fecha_Fin_Periodo As Date
Global StrSql2 As String

Global Sep As String


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

'    strCmdLine = Command()
'    ArrParametros = Split(strCmdLine, " ", -1)
'    If UBound(ArrParametros) > 0 Then
'        If IsNumeric(ArrParametros(0)) Then
'            NroProcesoBatch = ArrParametros(0)
'            Etiqueta = ArrParametros(1)
'        Else
'            Exit Sub
'        End If
'    Else
'        If IsNumeric(strCmdLine) Then
'            NroProcesoBatch = strCmdLine
'        Else
'            Exit Sub
'        End If
'    End If
    
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
'    OpenConnection strconexion, objConn
'    OpenConnection strconexion, objconnProgreso

        
    Nombre_Arch = PathFLog & "Exp_Jub_Mov_Excel" & "-" & NroProcesoBatch & ".log"
    
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
    
    'Cambio el estado del proceso a Procesando y el PID
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcprogreso = 0, bprcestado = 'Procesando', bprcpid = " & PID & " WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE btprcnro = 120 AND bpronro =" & NroProcesoBatch
    OpenRecordset StrSql, rs_batch_proceso
    
    TiempoInicialProceso = GetTickCount
    
    If Not rs_batch_proceso.EOF Then
        IdUser = rs_batch_proceso!IdUser
        Fecha = rs_batch_proceso!bprcfecha
        Hora = rs_batch_proceso!bprchora
        bprcparam = rs_batch_proceso!bprcparam
        
        rs_batch_proceso.Close
        Set rs_batch_proceso = Nothing
        
        Call LevantarParamteros(NroProcesoBatch, bprcparam)
    End If
    
    TiempoFinalProceso = GetTickCount
    Flog.writeline "Tiempo del proceso (milisegundos): " & (TiempoFinalProceso - TiempoInicialProceso)
    
    If Not HuboError Then
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado' WHERE bpronro = " & NroProcesoBatch
    Else
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Error' WHERE bpronro = " & NroProcesoBatch
    End If
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    Flog.Close
    objconnProgreso.Close
    objConn.Close
    
    Exit Sub
    
MainError:
    HuboError = True
    Flog.writeline " Error: " & Err.Description & Now
    
    
End Sub
Public Sub Generacion(ByVal FiltroEmpleado As String, ByVal bpronro As Long, ByVal Nroliq As Long, ByVal Todos_Pro As Boolean, ByVal Proc_Aprob As Integer, ByVal Empresa As Long, ByVal Agrupado As Boolean, _
    ByVal AgrupaTE1 As Boolean, ByVal Tenro1 As Long, Estrnro1 As Long, _
    ByVal AgrupaTE2 As Boolean, ByVal Tenro2 As Long, Estrnro2 As Long, _
    ByVal AgrupaTE3 As Boolean, ByVal Tenro3 As Long, Estrnro3 As Long)
' --------------------------------------------------------------------------------------------
' Descripcion: Procedimiento de generacion del Reporte de Declaracion Jurada de Aportes
' Autor      : FGZ
' Fecha      : 08/09/2004
' Ult. Mod   :
' Fecha      :
' --------------------------------------------------------------------------------------------
Dim Reg1 As TipoReg1
Dim Reg2 As TipoReg2
Dim Reg3 As TipoReg3

Dim Aux_Total_Importe
Dim strLinea As String
Dim Aux_Linea As String

Dim FechaDesde As Date
Dim FechaHasta As Date
Dim Nro_Reporte As Integer
Dim Conf_Ok As Boolean
Dim concnro As Long
Dim Nro_Concepto As Long

Dim Estructura1 As Long
Dim Estructura2 As Long

Dim rs_Confrep As New ADODB.Recordset
Dim rs_Concepto As New ADODB.Recordset
Dim rs_Detliq As New ADODB.Recordset
Dim rs_Doc As New ADODB.Recordset
Dim rs_Periodo As New ADODB.Recordset
Dim rs_Procesos As New ADODB.Recordset
Dim rs_Rep_jub_mov As New ADODB.Recordset
Dim rs_Estructura As New ADODB.Recordset
Dim rs_Modelo As New ADODB.Recordset

Const ForReading = 1
Const TristateFalse = 0
Dim fExport
Dim fAuxiliar
Dim Directorio As String
Dim Archivo As String
Dim Intentos As Integer
Dim carpeta

Dim Aux_str As String
Dim TipoCodEmpresa As String
Dim NroEmpresa As Long

'Archivo de exportacion
StrSql = "SELECT sis_dirsalidas FROM sistema WHERE sisnro = 1 "
OpenRecordset StrSql, rs
If Not rs.EOF Then
    Directorio = Trim(rs!sis_dirsalidas)
End If

StrSql = "SELECT * FROM modelo WHERE modnro = 228"
OpenRecordset StrSql, rs_Modelo
If Not rs_Modelo.EOF Then
    If Not IsNull(rs_Modelo!modarchdefault) Then
        Directorio = Directorio & Trim(rs_Modelo!modarchdefault)
    Else
        Flog.writeline "El modelo no tiene configurada la carpeta desteino. El archivo será generado en el directorio default"
    End If
Else
    Flog.writeline "No se encontró el modelo. El archivo será generado en el directorio default"
End If

'Obtengo los datos del separador
If Not rs_Modelo.EOF Then
   Sep = rs_Modelo!modseparador
Else
   Sep = ""
End If

Archivo = Directorio & "\jub-mov" & "-" & NroProcesoBatch & ".csv"
Set fs = CreateObject("Scripting.FileSystemObject")
'Activo el manejador de errores
On Error Resume Next
Set fExport = fs.CreateTextFile(Archivo, True)
If Err.Number <> 0 Then
    Flog.writeline "La carpeta Destino no existe. Se creará."
    Set carpeta = fs.CreateFolder(Directorio)
    Set fExport = fs.CreateTextFile(Archivo, True)
End If
'desactivo el manejador de errores
On Error GoTo CE

Archivo = Directorio & "\auxiliar.txt"
Set fs = CreateObject("Scripting.FileSystemObject")
Set fAuxiliar = fs.CreateTextFile(Archivo, True)

'cargo el periodo
StrSql = "SELECT * FROM periodo WHERE pliqnro = " & CStr(Nroliq)
OpenRecordset StrSql, rs_Periodo
If rs_Periodo.EOF Then
    Flog.writeline "No se encontró el Periodo"
    Exit Sub
End If
Fecha_Inicio_periodo = rs_Periodo!pliqdesde
Fecha_Fin_Periodo = rs_Periodo!pliqhasta

'Configuracion del Reporte
Nro_Reporte = 89
'Columna 1 - Tenro = 13 Agrupaciones (no es indispensable)
'Columna 2 - Tenro = 26 Categ CTT (no es indispensable)
'Columna 3 - Concepto (indispensable)
Conf_Ok = False
StrSql = "SELECT * FROM confrep WHERE repnro = " & Nro_Reporte
OpenRecordset StrSql, rs_Confrep
If rs_Confrep.EOF Then
    Flog.writeline "No se encontró la configuración del Reporte"
    Exit Sub
Else
    Do While Not rs_Confrep.EOF
        Select Case rs_Confrep!confnrocol
        Case 1:
            Estructura1 = rs_Confrep!confval
        Case 2:
            Estructura2 = rs_Confrep!confval
        Case 3:
            Nro_Concepto = rs_Confrep!confval
            StrSql = "SELECT * FROM concepto WHERE conccod = " & Nro_Concepto
            OpenRecordset StrSql, rs_Concepto
            If rs_Concepto.EOF Then
                Flog.writeline "Columna 1. El concepto no existe"
                Exit Sub
            Else
                Conf_Ok = True
                concnro = rs_Concepto!concnro
            End If
        End Select
        rs_Confrep.MoveNext
    Loop
End If
If Not Conf_Ok Then
    Flog.writeline "Columna 3. El concepto no esta configurado"
    Exit Sub
End If

' Comienzo la transaccion
MyBeginTrans

'Depuracion del Temporario
StrSql = "DELETE FROM rep_jub_mov "

objConn.Execute StrSql, , adExecuteNoRecords

StrSql = "SELECT empleado.* "
If AgrupaTE1 Then
    StrSql = StrSql & ", te1.tenro tenro1, te1.estrnro estrnro1"
End If
If AgrupaTE2 Then
    StrSql = StrSql & ", te2.tenro tenro2, te2.estrnro estrnro2"
End If
If AgrupaTE3 Then
    StrSql = StrSql & ", te3.tenro tenro3, te3.estrnro estrnro3"
End If
StrSql = StrSql & "  FROM  Empleado "
StrSql = StrSql & " INNER JOIN his_estructura  ON his_estructura.ternro = empleado.ternro and his_estructura.tenro = 10 "
StrSql = StrSql & " INNER JOIN empresa ON empresa.estrnro = his_estructura.estrnro AND empresa.estrnro =" & Empresa
If AgrupaTE1 Then
    StrSql = StrSql & " INNER JOIN his_estructura TE1 ON te1.ternro = empleado.ternro "
End If
If AgrupaTE2 Then
    StrSql = StrSql & " INNER JOIN his_estructura TE2 ON te2.ternro = empleado.ternro "
End If
If AgrupaTE3 Then
    StrSql = StrSql & " INNER JOIN his_estructura TE3 ON te3.ternro = empleado.ternro "
End If
StrSql = StrSql & " WHERE " & FiltroEmpleado
StrSql = StrSql & " AND empresa.estrnro =" & Empresa
StrSql = StrSql & " AND (his_estructura.htetdesde <= " & ConvFecha(Fecha_Fin_Periodo) & ") AND "
StrSql = StrSql & " ((" & ConvFecha(Fecha_Fin_Periodo) & " <= his_estructura.htethasta) or (his_estructura.htethasta is null))"
If AgrupaTE1 Then
    StrSql = StrSql & " AND  te1.tenro = " & Tenro1 & " AND "
    If Estrnro1 <> 0 Then
        StrSql = StrSql & " te1.estrnro = " & Estrnro1 & " AND "
    End If
    StrSql = StrSql & " (te1.htetdesde <= " & ConvFecha(Fecha_Fin_Periodo) & ") AND "
    StrSql = StrSql & " ((" & ConvFecha(Fecha_Fin_Periodo) & " <= te1.htethasta) or (te1.htethasta is null)) "
End If
If AgrupaTE2 Then
    StrSql = StrSql & " AND te2.tenro = " & Tenro2 & " AND "
    If Estrnro2 <> 0 Then
        StrSql = StrSql & " te2.estrnro = " & Estrnro2 & " AND "
    End If
    StrSql = StrSql & " (te2.htetdesde <= " & ConvFecha(Fecha_Fin_Periodo) & ") AND "
    StrSql = StrSql & " ((" & ConvFecha(Fecha_Fin_Periodo) & " <= te2.htethasta) or (te2.htethasta is null))  "
End If
If AgrupaTE3 Then
    StrSql = StrSql & " AND te3.tenro = " & Tenro3 & " AND "
    If Estrnro3 <> 0 Then
        StrSql = StrSql & " te3.estrnro = " & Estrnro3 & " AND "
    End If
    StrSql = StrSql & " (te3.htetdesde <= " & ConvFecha(Fecha_Fin_Periodo) & ") AND "
    StrSql = StrSql & " ((" & ConvFecha(Fecha_Fin_Periodo) & " <= te3.htethasta) or (te3.htethasta is null))"
End If
StrSql = StrSql & " ORDER BY empleado.ternro"
OpenRecordset StrSql, rs_Procesos

' Adrián - Busco el estrnro de la empresa
StrSql = "SELECT empnro, estrnro FROM empresa WHERE empresa.estrnro = " & Empresa
OpenRecordset StrSql, rs_Empresa

' Adrián - Busco el tipo de código Estrella de la empresa.
If rs_Empresa.EOF Then
            Flog.writeline " No existe una estructura para esta Empresa"
Else
    NroEmpresa = rs_Empresa!Empnro
    
   StrSql = "SELECT nrocod"
   StrSql = StrSql & " FROM estr_cod"
   StrSql = StrSql & " INNER JOIN tipocod ON tipocod.tcodnro = estr_cod.tcodnro"
   StrSql = StrSql & " WHERE (tipocod.tcodnro = 32)"
   StrSql = StrSql & " AND estrnro = " & rs_Empresa!estrnro
   OpenRecordset StrSql, rs_tipocod

    If rs_tipocod.EOF Then
        Flog.writeline " No existe número de La Estrella para esta Empresa"
        TipoCodEmpresa = String(15, "0")
    Else
        If Len(rs_tipocod!nrocod) < 15 Then
            TipoCodEmpresa = rs_tipocod!nrocod & String(15 - Len(rs_tipocod!nrocod), "0")
        Else
            TipoCodEmpresa = Left(rs_tipocod!nrocod, 15)
        End If
    End If
End If

'seteo de las variables de progreso
Progreso = 0
CConceptosAProc = rs_Procesos.RecordCount
If CConceptosAProc = 0 Then
    CConceptosAProc = 1
End If
IncPorc = (100 / CConceptosAProc)

'inicializo
Reg1.Tipo_Reg = "1"
'Adrián - Utilizo el nro de codigo Estrella para la empresa.
Reg1.Nro_ID = TipoCodEmpresa
Reg1.Total_Aportes = 0
Reg1.Salario_MesAno = Format(CStr(rs_Periodo!pliqmes), "00") & Right(CStr(rs_Periodo!pliqanio), 2)
Reg1.Total_Pag = "0001"
Reg1.Codigo_Declaracion = "1"

Reg2.Tipo_Reg = "2"
'Adrián - Utilizo el nro de codigo Estrella para la empresa.
Reg2.Nro_ID = TipoCodEmpresa
Reg2.Nro_Pag = "0001"
Reg2.Total_Aportes = 0
Reg2.Espacios = "      "
'Reg2.Espacios = "00000"

'seteo los valores que son fijos
Reg3.Tipo_Reg = "3"
'Adrián - Utilizo el nro de codigo Estrella para la empresa.
Reg3.Nro_ID = TipoCodEmpresa
Reg3.Espacios = " "
'Reg3.Espacios = "0"

Do While Not rs_Procesos.EOF
        Reg3.Importe = 0
        
        ' Buscar el documento
        StrSql = " SELECT ter_doc.tidnro, ter_doc.nrodoc FROM tercero " & _
                 " INNER JOIN ter_doc ON tercero.ternro = ter_doc.ternro " & _
                 " WHERE tercero.ternro= " & rs_Procesos!ternro & _
                 " ORDER BY ter_doc.tidnro "
        OpenRecordset StrSql, rs_Doc
        If Not rs_Doc.EOF Then
            Select Case rs_Doc!tidnro
            Case 1, 2, 3:
                Reg3.Tipo_Doc = "1"
            Case 4:
                Reg3.Tipo_Doc = "4"
            Case Else
                Reg3.Tipo_Doc = "1"
            End Select
            Reg3.Nro_Doc = Format_StrNro(Left(CStr(rs_Doc!nrodoc), 8), 8, True, "0")
        Else
            Flog.writeline "Error al obtener los datos del Documento"
            Reg3.Tipo_Doc = "1"
            Reg3.Nro_Doc = "00000000"
        End If
        
        'busco el concepto
'        StrSql = "SELECT * FROM detliq " & _
'                 " INNER JOIN cabliq ON detliq.cliqnro = cabliq.cliqnro " & _
'                 " INNER JOIN concepto ON detliq.concnro = concepto.concnro " & _
'                 " WHERE concepto.concnro = " & concnro & _
'                 " AND cabliq.cliqnro =" & rs_Procesos!cliqnro & _
'                 " AND (concepto.concimp = -1" & _
'                 " OR concepto.concpuente = 0)"
        StrSql = "SELECT detliq.* "
        StrSql = StrSql & "  FROM periodo "
        StrSql = StrSql & " INNER JOIN proceso ON periodo.pliqnro = proceso.pliqnro "
        StrSql = StrSql & " INNER JOIN cabliq ON proceso.pronro = cabliq.pronro "
        StrSql = StrSql & " INNER JOIN empleado ON cabliq.empleado = empleado.ternro "
        StrSql = StrSql & " INNER JOIN detliq ON detliq.cliqnro = cabliq.cliqnro "
        StrSql = StrSql & " INNER JOIN concepto ON detliq.concnro = concepto.concnro "
        StrSql = StrSql & " WHERE periodo.pliqnro =" & Nroliq
        If Not Todos_Pro Then
            StrSql = StrSql & " AND proceso.pronro IN (" & NroProc & ")"
        Else
            StrSql = StrSql & " AND proaprob = " & CInt(Proc_Aprob)
        End If
        StrSql = StrSql & " AND empleado.ternro =" & rs_Procesos!ternro
        StrSql = StrSql & " AND concepto.concnro = " & concnro
        StrSql = StrSql & " AND (concepto.concimp = -1"
        StrSql = StrSql & " OR concepto.concpuente = 0)"
        OpenRecordset StrSql, rs_Detliq
        Do While Not rs_Detliq.EOF
            'Adrián - Sumo el valor absoluto del monto.
            Reg3.Importe = Reg3.Importe + Abs(Round(rs_Detliq!dlimonto, 2))
            
            rs_Detliq.MoveNext
        Loop
        
    If Reg3.Importe <> 0 Then
        'Si no existe el rep_juv_mov
        StrSql = "SELECT * FROM rep_jub_mov "
        StrSql = StrSql & " WHERE ternro = " & rs_Procesos!ternro
        StrSql = StrSql & " AND bpronro = " & bpronro
        StrSql = StrSql & " AND pliqnro = " & Nroliq
        StrSql = StrSql & " AND empresa = " & NroEmpresa
        If Not Todos_Pro Then
            StrSql = StrSql & " AND pronro = '" & Left(ListaNroProc, 200) & "'"
        Else
            StrSql = StrSql & " AND pronro = '0'"
            StrSql = StrSql & " AND proaprob= " & CInt(Proc_Aprob)
        End If
        OpenRecordset StrSql, rs_Rep_jub_mov
    
        If rs_Rep_jub_mov.EOF Then
            'Inserto
            StrSql = "INSERT INTO rep_jub_mov (bpronro,pliqnro,pronro,proaprob,empresa,iduser,fecha,hora,"
            StrSql = StrSql & "tiporegistro,nroidentificador,tidnro,nrodoc,importe,"
            StrSql = StrSql & "ternro,empleg,apeynom,"
            StrSql = StrSql & "tenro1,estrnro1,tedesc1,estrdesc1,tenro2,estrnro2,tedesc2,estrdesc2,tenro3,estrnro3,tedesc3,estrdesc3 "
            StrSql = StrSql & ") VALUES ("
            StrSql = StrSql & bpronro & ","
            StrSql = StrSql & Nroliq & ","
            If Not Todos_Pro Then
                StrSql = StrSql & "'" & Left(ListaNroProc, 200) & "',"
                'StrSql = StrSql & rs_Procesos!proaprob & ","
                StrSql = StrSql & CInt(Proc_Aprob) & ","
            Else
                StrSql = StrSql & "0" & ","
                StrSql = StrSql & CInt(Proc_Aprob) & ","
            End If
            StrSql = StrSql & NroEmpresa & ","
            StrSql = StrSql & "'" & Left(IdUser, 20) & "',"
            StrSql = StrSql & ConvFecha(Fecha) & ","
            StrSql = StrSql & "'" & Left(Hora, 10) & "',"
            
            StrSql = StrSql & "'" & Left(Reg3.Tipo_Reg, 1) & "',"
            StrSql = StrSql & "'" & Left(Reg3.Nro_ID, 15) & "',"
            StrSql = StrSql & "'" & Left(Reg3.Tipo_Doc, 1) & "',"
            StrSql = StrSql & "'" & Left(Reg3.Nro_Doc, 8) & "',"
            StrSql = StrSql & CSng(Reg3.Importe) & ","
            
            StrSql = StrSql & rs_Procesos!ternro & ","
            StrSql = StrSql & rs_Procesos!empleg & ","
            
            'FGZ - 28/09/2004
            Aux_str = rs_Procesos!terape & IIf(Not IsNull(rs_Procesos!terape2), rs_Procesos!terape2, "")
            Aux_str = Aux_str & " " & rs_Procesos!ternom & IIf(Not IsNull(rs_Procesos!ternom2), rs_Procesos!ternom2, "")
            StrSql = StrSql & FormatearParaSql(Aux_str, 40, True, False)
            StrSql = StrSql & ","
            
            'Estructuras
            If AgrupaTE1 Then
                StrSql = StrSql & Tenro1 & ","
            Else
                StrSql = StrSql & "null" & ","
            End If
            StrSql = StrSql & Estrnro1 & ","
            
            'Descripcion tipo estructura
            If AgrupaTE1 Then
                StrSql2 = "SELECT * FROM tipoestructura WHERE tenro =" & rs_Procesos!Tenro1
                If rs_Estructura.State = adStateOpen Then rs_Estructura.Close
                OpenRecordset StrSql2, rs_Estructura
                If Not rs_Estructura.EOF Then
                    'StrSql = StrSql & "'" & rs_Estructura!tedabr & "'" & ","
                    StrSql = StrSql & FormatearParaSql(rs_Estructura!tedabr, 25, True, False) & ","
                Else
                    'StrSql = StrSql & "' '" & ","
                    StrSql = StrSql & FormatearParaSql(" ", 25, True, False) & ","
                End If
                'Descripcion Estructura
                StrSql2 = "SELECT * FROM estructura WHERE estrnro =" & rs_Procesos!Estrnro1
                If rs_Estructura.State = adStateOpen Then rs_Estructura.Close
                OpenRecordset StrSql2, rs_Estructura
                If Not rs_Estructura.EOF Then
                    'StrSql = StrSql & "'" & rs_Estructura!estrdabr & "'" & ","
                    StrSql = StrSql & FormatearParaSql(rs_Estructura!estrdabr, 25, True, False) & ","
                Else
                    'StrSql = StrSql & "' '" & ","
                    StrSql = StrSql & FormatearParaSql(" ", 25, True, False) & ","
                End If
            Else
                'StrSql = StrSql & "' '" & ","
                'StrSql = StrSql & "' '" & ","
                StrSql = StrSql & FormatearParaSql(" ", 25, True, False) & ","
                StrSql = StrSql & FormatearParaSql(" ", 25, True, False) & ","
            End If
            
            If AgrupaTE2 Then
                StrSql = StrSql & Tenro2 & ","
            Else
                StrSql = StrSql & "null" & ","
            End If
            StrSql = StrSql & Estrnro2 & ","
            
            If AgrupaTE2 Then
                'Descripcion tipo estructura
                StrSql2 = "SELECT * FROM tipoestructura WHERE tenro =" & rs_Procesos!Tenro2
                If rs_Estructura.State = adStateOpen Then rs_Estructura.Close
                OpenRecordset StrSql2, rs_Estructura
                If Not rs_Estructura.EOF Then
                    'StrSql = StrSql & "'" & rs_Estructura!tedabr & "'" & ","
                    StrSql = StrSql & FormatearParaSql(rs_Estructura!tedabr, 25, True, False) & ","
                Else
                    'StrSql = StrSql & "' '" & ","
                    StrSql = StrSql & FormatearParaSql(" ", 25, True, False) & ","
                End If
                'Descripcion Estructura
                StrSql2 = "SELECT * FROM estructura WHERE estrnro =" & rs_Procesos!Estrnro2
                If rs_Estructura.State = adStateOpen Then rs_Estructura.Close
                OpenRecordset StrSql2, rs_Estructura
                If Not rs_Estructura.EOF Then
                    'StrSql = StrSql & "'" & rs_Estructura!estrdabr & "'" & ","
                    StrSql = StrSql & FormatearParaSql(rs_Estructura!estrdabr, 25, True, False) & ","
                Else
                    'StrSql = StrSql & "' '" & ","
                    StrSql = StrSql & FormatearParaSql(" ", 25, True, False) & ","
                End If
            Else
                'StrSql = StrSql & "' '" & ","
                'StrSql = StrSql & "' '" & ","
                StrSql = StrSql & FormatearParaSql(" ", 25, True, False) & ","
                StrSql = StrSql & FormatearParaSql(" ", 25, True, False) & ","
            End If
            
            If AgrupaTE3 Then
                StrSql = StrSql & Tenro3 & ","
            Else
                StrSql = StrSql & "null" & ","
            End If
            StrSql = StrSql & Estrnro3 & ","
            
            'Descripcion tipo estructura
            If AgrupaTE3 Then
                StrSql2 = "SELECT * FROM tipoestructura WHERE tenro =" & rs_Procesos!Tenro3
                If rs_Estructura.State = adStateOpen Then rs_Estructura.Close
                OpenRecordset StrSql2, rs_Estructura
                If Not rs_Estructura.EOF Then
                    'StrSql = StrSql & "'" & rs_Estructura!tedabr & "'" & ","
                    StrSql = StrSql & FormatearParaSql(rs_Estructura!tedabr, 25, True, False) & ","
                Else
                    'StrSql = StrSql & "' '" & ","
                    StrSql = StrSql & FormatearParaSql(" ", 25, True, False) & ","
                End If
                'Descripcion Estructura
                StrSql2 = "SELECT * FROM estructura WHERE estrnro =" & rs_Procesos!Estrnro3
                If rs_Estructura.State = adStateOpen Then rs_Estructura.Close
                OpenRecordset StrSql2, rs_Estructura
                If Not rs_Estructura.EOF Then
                    'StrSql = StrSql & "'" & rs_Estructura!estrdabr & "'"
                    StrSql = StrSql & FormatearParaSql(rs_Estructura!estrdabr, 25, True, False)
                Else
                    'StrSql = StrSql & "' '"
                    StrSql = StrSql & FormatearParaSql(" ", 25, True, False)
                End If
            Else
                'StrSql = StrSql & "' '" & ","
                'StrSql = StrSql & "' '"
                StrSql = StrSql & FormatearParaSql(" ", 25, True, False) & ","
                StrSql = StrSql & FormatearParaSql(" ", 25, True, False)
            End If
            
            StrSql = StrSql & ")"
            objConn.Execute StrSql, , adExecuteNoRecords
        Else
            'Actualizo
            StrSql = "UPDATE rep_jub_mov SET importe = importe + " & Reg3.Importe
            StrSql = StrSql & " WHERE ternro = " & rs_Procesos!ternro
            StrSql = StrSql & " AND bpronro = " & bpronro
            StrSql = StrSql & " AND pliqnro = " & Nroliq
            StrSql = StrSql & " AND empresa = " & NroEmpresa
            If Not Todos_Pro Then
                StrSql = StrSql & " AND pronro = '" & Left(ListaNroProc, 200) & "'"
            Else
                StrSql = StrSql & " AND pronro = '0'"
                StrSql = StrSql & " AND proaprob= " & CInt(Proc_Aprob)
            End If
            objConn.Execute StrSql, , adExecuteNoRecords
        End If
    End If
    'Actualizo el progreso del Proceso
    Progreso = Progreso + IncPorc
    TiempoAcumulado = GetTickCount
    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & _
             ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & _
             "' WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
                
    If Reg3.Importe <> 0 Then
        Aux_Total_Importe = Format(Reg3.Importe, "0000000000000.00")
        Aux_Total_Importe = Replace(Aux_Total_Importe, ".", "")
        Aux_Total_Importe = Replace(Aux_Total_Importe, ",", "")
        Aux_Total_Importe = Format(Aux_Total_Importe, "000000000000000")
        If Len(Aux_Total_Importe) < 15 Then
            Aux_Total_Importe = String(15 - Len(Aux_Total_Importe), "0") & Aux_Total_Importe
        End If
        'Escribo en el auxiliar
        fAuxiliar.writeline Reg3.Tipo_Reg & Sep & Reg3.Nro_ID & Sep & Reg3.Tipo_Doc & Sep & Reg3.Nro_Doc & Sep & Aux_Total_Importe & Sep & Reg3.Espacios & Sep
    End If
    
    Reg1.Total_Aportes = Reg1.Total_Aportes + Reg3.Importe
    Reg2.Total_Aportes = Reg2.Total_Aportes + Reg3.Importe
    
    'Siguiente proceso
    rs_Procesos.MoveNext
Loop
fAuxiliar.Close

'Exportar archivo de texto
Aux_Total_Importe = Format(Reg1.Total_Aportes, "00000000000000.00")
Aux_Total_Importe = Replace(Aux_Total_Importe, ".", "")
Aux_Total_Importe = Replace(Aux_Total_Importe, ",", "")
Aux_Total_Importe = Format(Aux_Total_Importe, "0000000000000000")

Aux_Linea = Reg1.Tipo_Reg & Sep & Reg1.Nro_ID & Sep & Aux_Total_Importe & Sep & Reg1.Salario_MesAno & Sep & _
                Reg1.Total_Pag & Sep & Reg1.Codigo_Declaracion & Sep
fExport.writeline Aux_Linea
'fExport.Write Aux_Linea
Aux_Linea = Reg2.Tipo_Reg & Sep & Reg2.Nro_ID & Sep & Reg2.Nro_Pag & Sep & Aux_Total_Importe & Sep & Reg2.Espacios & Sep
fExport.writeline Aux_Linea
'fExport.Write Aux_Linea

'leo el auxiliar y lo escribo
    On Error Resume Next
    Intentos = 0
    Err.Number = 1
    Do Until Err.Number = 0 Or Intentos = 10
        Err.Number = 0
        Set fAuxiliar = fs.getfile(Archivo)
        If fAuxiliar.Size = 0 Then
            Err.Number = 1
            Intentos = Intentos + 1
        End If
    Loop
    On Error GoTo CE
   
   If Not Intentos = 10 Then
       'Abro el archivo
        On Error GoTo CE
        Set fAuxiliar = fs.OpenTextFile(Archivo, ForReading, TristateFalse)
    
        Do While Not fAuxiliar.AtEndOfStream
            strLinea = fAuxiliar.ReadLine
            fExport.writeline strLinea
            'fExport.Write strLinea
        Loop
        fAuxiliar.Close
    End If
    fExport.Close

    'Borro el auxiliar
    fs.Deletefile Archivo, True
    
'Fin de la transaccion
MyCommitTrans

If rs_Confrep.State = adStateOpen Then rs_Confrep.Close
If rs_Concepto.State = adStateOpen Then rs_Concepto.Close
If rs_Detliq.State = adStateOpen Then rs_Detliq.Close
If rs_Doc.State = adStateOpen Then rs_Doc.Close
If rs_Procesos.State = adStateOpen Then rs_Procesos.Close
If rs_Periodo.State = adStateOpen Then rs_Periodo.Close
If rs_Rep_jub_mov.State = adStateOpen Then rs_Rep_jub_mov.Close
If rs_Estructura.State = adStateOpen Then rs_Estructura.Close
If rs_tipocod.State = adStateOpen Then rs_tipocod.Close
If rs_Empresa.State = adStateOpen Then rs_Empresa.Close

Set rs_Confrep = Nothing
Set rs_Concepto = Nothing
Set rs_Detliq = Nothing
Set rs_Doc = Nothing
Set rs_Procesos = Nothing
Set rs_Periodo = Nothing
Set rs_Rep_jub_mov = Nothing
Set rs_Estructura = Nothing
Set rs_tipocod = Nothing
Set rs_Empresa = Nothing

Exit Sub
CE:
    HuboError = True
    MyRollbackTrans
    
    Flog.writeline " Error: " & Err.Description
    Flog.writeline " Ultima sql ejecutada: " & StrSql

    If rs_Confrep.State = adStateOpen Then rs_Confrep.Close
    If rs_Concepto.State = adStateOpen Then rs_Concepto.Close
    If rs_Detliq.State = adStateOpen Then rs_Detliq.Close
    If rs_Doc.State = adStateOpen Then rs_Doc.Close
    If rs_Procesos.State = adStateOpen Then rs_Procesos.Close
    If rs_Periodo.State = adStateOpen Then rs_Periodo.Close
    If rs_Rep_jub_mov.State = adStateOpen Then rs_Rep_jub_mov.Close
    If rs_Estructura.State = adStateOpen Then rs_Estructura.Close
    If rs_tipocod.State = adStateOpen Then rs_tipocod.Close
    If rs_Empresa.State = adStateOpen Then rs_Empresa.Close

    
    Set rs_Confrep = Nothing
    Set rs_Concepto = Nothing
    Set rs_Detliq = Nothing
    Set rs_Doc = Nothing
    Set rs_Procesos = Nothing
    Set rs_Periodo = Nothing
    Set rs_Rep_jub_mov = Nothing
    Set rs_Estructura = Nothing
    Set rs_tipocod = Nothing
    Set rs_Empresa = Nothing

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
'filtro de empleados
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
        FiltroEmpleados = Mid(parametros, pos1, pos2 - pos1 + 1)
    
        pos1 = pos2 + 2
        pos2 = InStr(pos1, parametros, Separador) - 1
        pliqdesde = CLng(Mid(parametros, pos1, pos2 - pos1 + 1))
    
        pos1 = pos2 + 2
        pos2 = InStr(pos1, parametros, Separador) - 1
        pliqhasta = CLng(Mid(parametros, pos1, pos2 - pos1 + 1))
    
        pos1 = pos2 + 2
        pos2 = InStr(pos1, parametros, Separador) - 1
        FechaDesde = CDate(Mid(parametros, pos1, pos2 - pos1 + 1))
        
        pos1 = pos2 + 2
        pos2 = InStr(pos1, parametros, Separador) - 1
        FechaHasta = CDate(Mid(parametros, pos1, pos2 - pos1 + 1))
    
        pos1 = pos2 + 2
        pos2 = InStr(pos1, parametros, Separador) - 1
        Proc_Aprob = Mid(parametros, pos1, pos2 - pos1 + 1)
    
        pos1 = pos2 + 2
        pos2 = InStr(pos1, parametros, Separador) - 1
        NroProc = Mid(parametros, pos1, pos2 - pos1 + 1)
        If NroProc = "0" Then
            Todos_Pro = True
        Else
            Todos_Pro = False
        End If
        ListaNroProc = Replace(NroProc, ",", "-")
        
        pos1 = pos2 + 2
        pos2 = InStr(pos1, parametros, Separador) - 1
        Tenro1 = Mid(parametros, pos1, pos2 - pos1 + 1)
        If Not Tenro1 = 0 Then
            Agrupado = True
            AgrupaTE1 = True
        End If
        pos1 = pos2 + 2
        pos2 = InStr(pos1, parametros, Separador) - 1
        Estrnro1 = Mid(parametros, pos1, pos2 - pos1 + 1)
        
        pos1 = pos2 + 2
        pos2 = InStr(pos1, parametros, Separador) - 1
        Tenro2 = Mid(parametros, pos1, pos2 - pos1 + 1)
        If Not Tenro2 = 0 Then
            AgrupaTE2 = True
        End If
        pos1 = pos2 + 2
        pos2 = InStr(pos1, parametros, Separador) - 1
        Estrnro2 = Mid(parametros, pos1, pos2 - pos1 + 1)
        
        pos1 = pos2 + 2
        pos2 = InStr(pos1, parametros, Separador) - 1
        Tenro3 = Mid(parametros, pos1, pos2 - pos1 + 1)
        If Not Tenro3 = 0 Then
            AgrupaTE3 = True
        End If
        pos1 = pos2 + 2
        pos2 = InStr(pos1, parametros, Separador) - 1
        Estrnro3 = Mid(parametros, pos1, pos2 - pos1 + 1)
        
        pos1 = pos2 + 2
        pos2 = InStr(pos1, parametros, Separador) - 1
        Empresa = Mid(parametros, pos1, pos2 - pos1 + 1)
    End If
End If

Call Generacion(FiltroEmpleados, bpronro, pliqdesde, Todos_Pro, Proc_Aprob, Empresa, Agrupado, AgrupaTE1, Tenro1, Estrnro1, AgrupaTE2, Tenro2, Estrnro2, AgrupaTE3, Tenro3, Estrnro3)

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

