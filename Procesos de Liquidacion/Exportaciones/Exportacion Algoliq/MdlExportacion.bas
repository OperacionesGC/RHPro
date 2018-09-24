Attribute VB_Name = "MdlExportacion"
Option Explicit

Private Type TR_Datos_Varios
    Convenio_Lecop As String        'String   long 4  -
    Filler As String                'String   long 1  -
    Cliente_Ya_Existente As String  'String   long 1  -
End Type

Global IdUser As String
Global Fecha As Date
Global Hora As String

Global Fecha_Inicio_periodo As Date
Global Fecha_Fin_Periodo As Date
Global StrSql2 As String

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
    
    ' carga las configuraciones basicas, formato de fecha, string de conexion,
    ' tipo de BD y ubicacion del archivo de log
    Call CargarConfiguracionesBasicas

    'Abro la conexion
    OpenConnection strconexion, objConn
    OpenConnection strconexion, objconnProgreso
    
    Nombre_Arch = PathFLog & "Exportacion" & "-" & NroProcesoBatch & ".log"
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Nombre_Arch, True)
    
    ' Obtengo el Process ID
    PID = GetCurrentProcessId
    Flog.Writeline Espacios(Tabulador * 0) & "PID = " & PID
    'Cambio el estado del proceso a Procesando y el PID
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcprogreso = 0, bprcestado = 'Procesando', bprcpid = " & PID & " WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE btprcnro = 61 AND bpronro =" & NroProcesoBatch
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
    Flog.Writeline
    Flog.Writeline Espacios(Tabulador * 0) & "=================================================="
    Flog.Writeline Espacios(Tabulador * 0) & "Tiempo del proceso (milisegundos): " & (TiempoFinalProceso - TiempoInicialProceso)
    Flog.Writeline Espacios(Tabulador * 0) & "=================================================="
    
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

Public Sub Generacion(ByVal bpronro As Long, ByVal Proceso As Long)
' --------------------------------------------------------------------------------------------
' Descripcion:  Procedimiento de generacion de los archivos
'               Exporta 4 Conceptos de Producci¢n para ALGOLIQ:
'               Conceptos: 290, 291, 292, 293. Cantidad y Monto
'               Exporta un archivo por empaque
'                   /producci¢n-remp00.txt   VILLA REGINA
'                   /producci¢n-nemp00.txt   VISTA ALEGRE
'              Formato Generado:
'                   Empleado 1-5
'                   Concepto 6-8
'                   Unidades 9-17 (9 Posiciones, 7 enteros y 2 decimales)
'                   Monto    18-26(9 Posiciones, 7 enteros y 2 decimales)
'                   Orden    27-27 (secuenciador entre novedades del mismo empleado)
'                   Fecha    28-35 (yyyymmdd)
' Autor      : FGZ
' Fecha      : 12/11/2004
' Ult. Mod   :
' Fecha      :
' --------------------------------------------------------------------------------------------
Const ForReading = 1
Const TristateFalse = 0

Dim fExport1
Dim fExport2
Dim fAuxiliar
Dim Directorio As String
Dim Archivo1 As String
Dim Archivo2 As String
Dim Intentos As Integer
Dim carpeta

Dim strLinea As String
Dim Aux_Linea As String

Dim Cantidad As Long
Dim Empaque As String
Dim Conccod1 As String
Dim Conccod2 As String
Dim Conccod3 As String
Dim Conccod4 As String

Dim Legajo As String
Dim concepto As String
Dim Unidades As String
Dim Monto As String
Dim Orden As String
Dim Fecha As String

'Registros
Dim rs_Cabliq As New ADODB.Recordset
Dim rs_Detliq As New ADODB.Recordset
Dim rs_Sucursal As New ADODB.Recordset
Dim rs_Modelo As New ADODB.Recordset


'Conceptos fijos
Conccod1 = "pr290"
Conccod2 = "pr291"
Conccod3 = "pr292"
Conccod4 = "pr293"

'para probar
'Conccod1 = "01100"
'Conccod2 = "01000"
'Conccod3 = "06005"
'Conccod4 = "06010"


'Archivo de exportacion
StrSql = "SELECT sis_dirsalidas FROM sistema WHERE sisnro = 1 "
OpenRecordset StrSql, rs
If Not rs.EOF Then
    Directorio = Trim(rs!sis_dirsalidas)
End If

StrSql = "SELECT * FROM modelo WHERE modnro = 238"
OpenRecordset StrSql, rs_Modelo
If Not rs_Modelo.EOF Then
    If Not IsNull(rs_Modelo!modarchdefault) Then
        Directorio = Directorio & Trim(rs_Modelo!modarchdefault)
    Else
        Flog.Writeline Espacios(Tabulador * 1) & "El modelo no tiene configurada la carpeta desteino. El archivo será generado en el directorio default"
    End If
Else
    Flog.Writeline Espacios(Tabulador * 1) & "No se encontró el modelo. El archivo será generado en el directorio default"
End If


'Seteo el nombre del archivo generado
Archivo1 = Directorio & "\produccion-remp00" & ".txt"   'VILLA REGINA
Archivo2 = Directorio & "\produccion-nemp00" & ".txt"   'VISTA ALEGRE
Set fs = CreateObject("Scripting.FileSystemObject")
'Activo el manejador de errores
On Error Resume Next
Set fExport1 = fs.CreateTextFile(Archivo1, True)
Set fExport2 = fs.CreateTextFile(Archivo2, True)
If Err.Number <> 0 Then
    Flog.Writeline Espacios(Tabulador * 1) & "La carpeta Destino no existe. Se creará."
    Set carpeta = fs.CreateFolder(Directorio)
    Set fExport1 = fs.CreateTextFile(Archivo1, True)
    Set fExport1 = fs.CreateTextFile(Archivo2, True)
End If
'desactivo el manejador de errores
On Error GoTo 0

'Comienzo la transaccion
MyBeginTrans

'Busco los procesos a evaluar
StrSql = "SELECT * FROM cabliq "
StrSql = StrSql & " INNER JOIN empleado ON empleado.ternro = cabliq.empleado "
StrSql = StrSql & " WHERE cabliq.pronro =" & Proceso
OpenRecordset StrSql, rs_Cabliq

'seteo de las variables de progreso
Progreso = 0
CConceptosAProc = rs_Cabliq.RecordCount
If CConceptosAProc = 0 Then
    CConceptosAProc = 1
    Flog.Writeline Espacios(Tabulador * 1) & "No hay nada para procesar "
End If
IncPorc = (100 / CConceptosAProc)

'Procesamiento
If rs_Cabliq.EOF Then
    Flog.Writeline Espacios(Tabulador * 1) & "No hay nada que procesar"
Else
    Flog.Writeline Espacios(Tabulador * 1) & "-------------------------------------"
    Flog.Writeline Espacios(Tabulador * 1) & "Exportando ..."
    Flog.Writeline
End If

Do While Not rs_Cabliq.EOF
        Cantidad_Warnings = 0
        
        Legajo = Format(rs_Cabliq!empleg, "00000")
        
        StrSql = "SELECT * FROM his_estructura "
        StrSql = StrSql & " INNER JOIN estructura ON his_estructura.estrnro = estructura.estrnro "
        StrSql = StrSql & " WHERE his_estructura.ternro =" & rs_Cabliq!ternro
        OpenRecordset StrSql, rs_Sucursal
        If Not rs_Sucursal.EOF Then
            If UCase(rs_Sucursal!estrcodext) = "1" Then
                Empaque = "REMP00"
            Else
                Empaque = "NEMP00"
            End If
        End If
        
        StrSql = "SELECT * FROM detliq "
        StrSql = StrSql & " INNER JOIN concepto ON detliq.concnro = concepto.concnro "
        StrSql = StrSql & " WHERE detliq.cliqnro =" & rs_Cabliq!cliqnro
        StrSql = StrSql & " AND ("
        StrSql = StrSql & " concepto.conccod = '" & Conccod1 & "'"
        StrSql = StrSql & " OR concepto.conccod = '" & Conccod2 & "'"
        StrSql = StrSql & " OR concepto.conccod = '" & Conccod3 & "'"
        StrSql = StrSql & " OR concepto.conccod = '" & Conccod4 & "'"
        StrSql = StrSql & " )"
        OpenRecordset StrSql, rs_Detliq
        
        Cantidad = 0
        Do While Not rs_Detliq.EOF
            Cantidad = Cantidad + 1
                    
            concepto = Right(rs_Detliq!Conccod, 3)
            
            Unidades = Format(rs_Detliq!dlicant, "00000000.00")
            Unidades = Mid(Unidades, 2, 7) & Right(Unidades, 2)
            
            'se supone que los monto son positivos
            Monto = Format(rs_Detliq!dlimonto, "00000000.00")
            Monto = Mid(Monto, 2, 7) & Right(Monto, 2)
            
            Orden = Format(Cantidad, "0")
            
            Fecha = Format(CDate(Now), "yyyymmdd")
        
            strLinea = Legajo & concepto & Unidades & Monto & Orden & Fecha
        
            ' ------------------------------------------------------------------------
            'Escribo en el archivo de texto
            If Empaque = "REMP00" Then
                fExport1.Writeline strLinea
            Else
                fExport2.Writeline strLinea
            End If
            
           rs_Detliq.MoveNext
        Loop
        
        'Actualizo el progreso del Proceso
        Progreso = Progreso + IncPorc
        TiempoAcumulado = GetTickCount
        StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & _
                 ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & _
                 "' WHERE bpronro = " & NroProcesoBatch
        objconnProgreso.Execute StrSql, , adExecuteNoRecords
        
    'Siguiente cabecera
    rs_Cabliq.MoveNext
Loop
'Cierro el archivo creado
fExport1.Close
fExport2.Close

'Fin de la transaccion
MyCommitTrans


If rs_Cabliq.State = adStateOpen Then rs_Cabliq.Close

Set rs_Cabliq = Nothing

Exit Sub
CE:
    HuboError = True
    MyRollbackTrans

    If rs_Cabliq.State = adStateOpen Then rs_Cabliq.Close
    
    Set rs_Cabliq = Nothing
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

Dim Proceso As Long
'Orden de los parametros
'pronro

Separador = "@"
' Levanto cada parametro por separado
If Not IsNull(parametros) Then
    If Len(parametros) >= 1 Then
        pos1 = 1
        pos2 = Len(parametros)
        Proceso = Mid(parametros, pos1, pos2 - pos1 + 1)
        
'        pos1 = pos2 + 2
'        pos2 = Len(parametros)
'        Asiento = CDate(Mid(parametros, pos1, pos2 - pos1 + 1))
        
'        pos1 = pos2 + 2
'        pos2 = Len(parametros)
'        Empresa = Mid(parametros, pos1, pos2 - pos1 + 1)
    End If
End If
Call Generacion(bpronro, Proceso)
End Sub

