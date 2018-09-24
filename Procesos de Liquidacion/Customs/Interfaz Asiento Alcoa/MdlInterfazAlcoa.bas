Attribute VB_Name = "MdInterfazAlcoa"
Option Explicit

'Const Version = 1.01
'Const FechaVersion = "13/04/2007"   'Martin Ferraro - Version Inicial

'Const Version = 1.02
'Const FechaVersion = "18/04/2007"   'Martin Ferraro - Quitar caracteres "/", "~", "^", "`"
''                                                     El caracter de Haber debe ser C

'Const Version = 1.03
'Const FechaVersion = "23/04/2007"   'Martin Ferraro - Agrupacion de lineas

Global Const Version = "1.04" ' Cesar Stankunas
Global Const FechaVersion = "05/08/2009"
Global Const UltimaModificacion = ""    'Encriptacion de string connection

Global IdUser As String
Global Fecha As Date
Global Hora As String


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
    
    Nombre_Arch = PathFLog & "Exp_Detalle_Asiento_Contable" & "-" & NroProcesoBatch & ".log"
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Nombre_Arch, True)
    
    ' Obtengo el Process ID
    PID = GetCurrentProcessId
    Flog.writeline "-------------------------------------------------"
    Flog.writeline "-------------------------------------------------"
    Flog.writeline "Version                  : " & Version
    Flog.writeline "Fecha Ultima Modificacion: " & FechaVersion
    Flog.writeline "PID                      : " & PID
    Flog.writeline "-------------------------------------------------"
    Flog.writeline "-------------------------------------------------"
    Flog.writeline
    
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
    
    Flog.writeline
    
    'Cambio el estado del proceso a Procesando y el PID
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcprogreso = 0, bprcestado = 'Procesando', bprcpid = " & PID & " WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE btprcnro = 165 AND bpronro =" & NroProcesoBatch
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

Public Sub Generacion(ByVal bpronro As Long, ByVal Vol_Cod As Long)
' --------------------------------------------------------------------------------------------
' Descripcion: Procedimiento de generacion del archivo de Detalle del Asiento Contable
' Autor      : FGZ
' Fecha      : 16/12/2004
' Ult. Mod   : 24/08/2006 - Se agrego desc del modelo
' Fecha      :
' --------------------------------------------------------------------------------------------
Const ForReading = 1
Const TristateFalse = 0

Dim fExport
Dim fAuxiliar
Dim Directorio As String
Dim Archivo As String
Dim Intentos As Integer
Dim carpeta

Dim NroLiq As Long
Dim strLinea As String
Dim Aux_Linea As String
Dim Texto As String
Dim montoDetaTotal As Double
Dim montolineaTotal As Double
Dim montoDetaTotalStr As String
Dim hayDatos As Boolean
Dim canRegs As Long

Dim CorteD_H As Integer
Dim CorteCuenta As String
Dim CorteDesc As String
Dim CorteLinadesc As String
Dim ImprimioLinea As Boolean


'Registros
Dim rs_Modelo As New ADODB.Recordset
Dim rs_Detalles As New ADODB.Recordset
Dim rs_Periodo As New ADODB.Recordset

Dim fechaAsiento As String

Flog.writeline Espacios(Tabulador * 0) & "===================================================="
Flog.writeline Espacios(Tabulador * 0) & "Comienzo de la exportación "
Flog.writeline Espacios(Tabulador * 0) & "===================================================="
Flog.writeline

Flog.writeline Espacios(Tabulador * 0) & "Buscando Directorio de salida del sistema."
'Archivo de exportacion
StrSql = "SELECT sis_dirsalidas FROM sistema WHERE sisnro = 1 "
OpenRecordset StrSql, rs
If Not rs.EOF Then
    Directorio = Trim(rs!sis_dirsalidas)
    Flog.writeline Espacios(Tabulador * 1) & "Directorio de salida del sistema: " & Directorio
Else
    Flog.writeline Espacios(Tabulador * 1) & "ERROR. No se encontro Directorio de salida del sistema."
    Exit Sub
End If
Flog.writeline

Flog.writeline Espacios(Tabulador * 0) & "Buscando directorio configurado en el modelo 234."
StrSql = "SELECT * FROM modelo WHERE modnro = 234"
OpenRecordset StrSql, rs_Modelo
If Not rs_Modelo.EOF Then
    If Not IsNull(rs_Modelo!modarchdefault) Then
        Directorio = Directorio & Trim(rs_Modelo!modarchdefault)
        Flog.writeline Espacios(Tabulador * 1) & "El archivo se genera en: " & Directorio
    Else
        Flog.writeline Espacios(Tabulador * 1) & "El modelo no tiene configurada la carpeta desteino. El archivo será generado en el directorio default"
    End If
Else
    Flog.writeline Espacios(Tabulador * 1) & "No se encontró el modelo. El archivo será generado en el directorio default"
End If
Flog.writeline

'cargo el periodo
Flog.writeline Espacios(Tabulador * 0) & "Buscando el proceso de Volcado."
StrSql = "SELECT * FROM proc_vol "
StrSql = StrSql & " WHERE vol_cod = " & Vol_Cod
OpenRecordset StrSql, rs_Periodo
If rs_Periodo.EOF Then
    Flog.writeline Espacios(Tabulador * 1) & "ERROR. No se encontro Proceso de Volcado."
    Exit Sub
Else
    Flog.writeline Espacios(Tabulador * 1) & "Proceso de Volcado: " & Vol_Cod & " " & rs_Periodo!vol_desc
    fechaAsiento = rs_Periodo!vol_fec_asiento
End If
Flog.writeline

'Seteo el nombre del archivo generado
Archivo = Directorio & "\Exp_" & rs_Periodo!vol_desc & "_" & Format(fechaAsiento, "ddmmyyyy") & ".txt"
Set fs = CreateObject("Scripting.FileSystemObject")
'Activo el manejador de errores
On Error Resume Next
Set fExport = fs.CreateTextFile(Archivo, True)
If Err.Number <> 0 Then
    Flog.writeline Espacios(Tabulador * 1) & "La carpeta Destino no existe. Se creará."
    Set carpeta = fs.CreateFolder(Directorio)
    Set fExport = fs.CreateTextFile(Archivo, True)
End If
'desactivo el manejador de errores
On Error GoTo 0

' Comienzo la transaccion
'MyBeginTrans

Flog.writeline Espacios(Tabulador * 0) & "Buscando Detalles a procesar."
'Busco los procesos a evaluar
StrSql = "SELECT detalle_asi.*, mod_asiento.masidesc  FROM  detalle_asi "
StrSql = StrSql & " INNER JOIN mod_asiento ON mod_asiento.masinro = detalle_asi.masinro "
StrSql = StrSql & " WHERE detalle_asi.vol_cod =" & Vol_Cod
StrSql = StrSql & " ORDER BY detalle_asi.linadesc, detalle_asi.linaD_H, detalle_asi.cuenta, detalle_asi.dldescripcion "
OpenRecordset StrSql, rs_Detalles

'seteo de las variables de progreso
Progreso = 0
CConceptosAProc = rs_Detalles.RecordCount
If CConceptosAProc = 0 Then
    CConceptosAProc = 1
End If
IncPorc = (100 / CConceptosAProc)

hayDatos = True

'Procesamiento
If rs_Detalles.EOF Then
    Flog.writeline Espacios(Tabulador * 1) & "No hay Detalles para ese Proceso de Volcado " & Vol_Cod
    hayDatos = False
Else
    'Inicializo las varables de corte
    CorteD_H = rs_Detalles!linaD_H
    CorteCuenta = rs_Detalles!Cuenta
    CorteDesc = rs_Detalles!dldescripcion
    CorteLinadesc = rs_Detalles!Linadesc
    ImprimioLinea = False
End If

montoDetaTotal = 0
montolineaTotal = 0
canRegs = 0
Do While Not rs_Detalles.EOF
        
        Flog.writeline Espacios(Tabulador * 1) & "Procesando detalle numero: " & rs_Detalles!detasinro
        
        'Miro si cambio de grupo
        If ((CorteD_H <> rs_Detalles!linaD_H) Or (CorteCuenta <> rs_Detalles!Cuenta) Or (CorteDesc <> rs_Detalles!dldescripcion)) Then
                    
            'CAMBIO DE REGISTRO - DEBO EXPORTAR LOS DATOS
            
            'Escribo en el archivo de texto
            Aux_Linea = EscribirRegistro(fechaAsiento, CorteD_H, montolineaTotal, CorteLinadesc, CorteCuenta, CorteDesc)
            fExport.writeline Aux_Linea
            
            'Actualizo la cantidad de registros
            canRegs = canRegs + 1
        
            'Actualizo las variables de linea
            CorteD_H = rs_Detalles!linaD_H
            CorteCuenta = rs_Detalles!Cuenta
            CorteDesc = rs_Detalles!dldescripcion
            CorteLinadesc = rs_Detalles!Linadesc
            'Total de asiento
            montoDetaTotal = montoDetaTotal + IIf(EsNulo(montolineaTotal), 0, montolineaTotal)
            'Total de la linea
            montolineaTotal = IIf(EsNulo(rs_Detalles!dlmonto), 0, rs_Detalles!dlmonto)
        
        Else
            montolineaTotal = montolineaTotal + IIf(EsNulo(rs_Detalles!dlmonto), 0, rs_Detalles!dlmonto)
            'CorteLinadesc = rs_Detalles!Linadesc
        
        End If
            
        'Actualizo el progreso del Proceso
        Progreso = Progreso + IncPorc
        TiempoAcumulado = GetTickCount
        StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & _
                 ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & _
                 "' WHERE bpronro = " & NroProcesoBatch
        objconnProgreso.Execute StrSql, , adExecuteNoRecords
                
    'Siguiente proceso
    rs_Detalles.MoveNext
Loop

'Imprimo registro de pie
If hayDatos Then
    
    'Para el caso del ultimo registro
    canRegs = canRegs + 1
    montoDetaTotal = montoDetaTotal + IIf(EsNulo(montolineaTotal), 0, montolineaTotal)
    
    'Escribo en el archivo de texto
    Aux_Linea = EscribirRegistro(fechaAsiento, CorteD_H, montolineaTotal, CorteLinadesc, CorteCuenta, CorteDesc)
    fExport.writeline Aux_Linea
    
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 0) & "Imprimiendo Pie."
    Flog.writeline Espacios(Tabulador * 1) & "Cant Reg: " & canRegs
    Flog.writeline Espacios(Tabulador * 1) & "Monto Total: " & Round(montoDetaTotal, 3)
    'Location - Fijo 408
    Aux_Linea = "408"
    
    'Accouting Period
    Aux_Linea = Aux_Linea & IIf(EsNulo(fechaAsiento), "0000", Format(fechaAsiento, "yymm"))
    
    'Journal Name
    Aux_Linea = Aux_Linea & Format_Str("408-FO-01", 20, True, " ")
    
    'Record Type
    Aux_Linea = Aux_Linea & "BR"
    
    'Batch Date
    Aux_Linea = Aux_Linea & IIf(EsNulo(fechaAsiento), "000000", Format(fechaAsiento, "yymmdd"))
    
    'Accounted Currency Control Amt - Fijo 19 ""
    Aux_Linea = Aux_Linea & Format_Str("", 19, True, " ")
    
    'Entered Currency Control Amt - Monto total
    montoDetaTotalStr = CStr(Round(montoDetaTotal, 3))
    montoDetaTotalStr = Format(montoDetaTotalStr, "##############0.000")
    'Aux_Linea = Aux_Linea & montoDetaTotalStr
    Aux_Linea = Aux_Linea & Format_StrNro(montoDetaTotalStr, 19, True, " ")
    
    'Statiscal Control Amount - Fijo 22 ""
    Aux_Linea = Aux_Linea & Format_Str("", 22, True, " ")
    
    'Record Count
    Aux_Linea = Aux_Linea & Format_StrNro(canRegs, 6, True, "0")
    
    'Filler - Fijo 899 ""
    Aux_Linea = Aux_Linea & Format_Str("", 899, True, " ")
    
    fExport.writeline Aux_Linea
    
End If


'Cierro el archivo creado
fExport.Close

'Fin de la transaccion
'MyCommitTrans


Fin:
If rs_Detalles.State = adStateOpen Then rs_Detalles.Close
If rs_Modelo.State = adStateOpen Then rs_Modelo.Close
If rs_Periodo.State = adStateOpen Then rs_Periodo.Close

Set rs_Detalles = Nothing
Set rs_Modelo = Nothing
Set rs_Periodo = Nothing

Exit Sub
CE:
    HuboError = True
    'MyRollbackTrans
    GoTo Fin
End Sub

Public Sub LevantarParamteros(ByVal bpronro As Long, ByVal parametros As String)
' --------------------------------------------------------------------------------------------
' Descripcion: Procedimiento para levantar los parametros pasados en batch_proceso en bprcparam
' Autor      : Martin Ferraro
' Fecha      :
' Ult. Mod   :
' Fecha      :
' --------------------------------------------------------------------------------------------
Dim Vol_Cod As Long
Dim Periodo As Long
Dim aux As String

'Orden de los parametros
'proceso de volcado

'Levanto cada parametro por separado
Flog.writeline Espacios(Tabulador * 0) & "Buscando Parametros."
If Not IsNull(parametros) Then
    If Len(parametros) >= 1 Then
        Vol_Cod = parametros
    Else
        Flog.writeline Espacios(Tabulador * 1) & "ERROR. Faltan Parametros."
        Exit Sub
    End If
Else
    Flog.writeline Espacios(Tabulador * 1) & "ERROR. Faltan Parametros."
    Exit Sub
End If

Flog.writeline Espacios(Tabulador * 1) & "Proceso de Volcado = " & Vol_Cod
Flog.writeline
Call Generacion(bpronro, Vol_Cod)
End Sub


Public Function QuitarCaracteresErroneos(ByVal cadena As String) As String
' --------------------------------------------------------------------------------------------
' Descripcion: Devuelve el string cadena sin los caracteres "/", "~", "^", "`"
' Autor      : Martin Ferraro
' Fecha      :
' Ult. Mod   :
' Fecha      :
' --------------------------------------------------------------------------------------------
Dim I As Long
Dim CadenaAux As String
    
    CadenaAux = cadena
    
    'Analisis de "/"
    Do While InStr(CadenaAux, "/") > 0
    
        'Primer caracter
        If InStr(CadenaAux, "/") = 1 Then
             CadenaAux = Right(CadenaAux, Len(CadenaAux) - 1)
        Else
            'Esta al medio
            If InStr(CadenaAux, "/") <> Len(CadenaAux) Then
                CadenaAux = Mid(CadenaAux, 1, InStr(CadenaAux, "/") - 1) & Mid(CadenaAux, InStr(CadenaAux, "/") + 1, Len(CadenaAux))
            Else
            'Es el ultimo
                CadenaAux = Left(CadenaAux, Len(CadenaAux) - 1)
            End If
        End If
        
    Loop
    
    'Analisis de "~"
    Do While InStr(CadenaAux, "~") > 0
    
        'Primer caracter
        If InStr(CadenaAux, "~") = 1 Then
             CadenaAux = Right(CadenaAux, Len(CadenaAux) - 1)
        Else
            'Esta al medio
            If InStr(CadenaAux, "~") <> Len(CadenaAux) Then
                CadenaAux = Mid(CadenaAux, 1, InStr(CadenaAux, "~") - 1) & Mid(CadenaAux, InStr(CadenaAux, "~") + 1, Len(CadenaAux))
            Else
            'Es el ultimo
                CadenaAux = Left(CadenaAux, Len(CadenaAux) - 1)
            End If
        End If
        
    Loop
    
    'Analisis de "^"
    Do While InStr(CadenaAux, "^") > 0
    
        'Primer caracter
        If InStr(CadenaAux, "^") = 1 Then
             CadenaAux = Right(CadenaAux, Len(CadenaAux) - 1)
        Else
            'Esta al medio
            If InStr(CadenaAux, "^") <> Len(CadenaAux) Then
                CadenaAux = Mid(CadenaAux, 1, InStr(CadenaAux, "^") - 1) & Mid(CadenaAux, InStr(CadenaAux, "^") + 1, Len(CadenaAux))
            Else
            'Es el ultimo
                CadenaAux = Left(CadenaAux, Len(CadenaAux) - 1)
            End If
        End If
        
    Loop
    
    'Analisis de "`"
    Do While InStr(CadenaAux, "`") > 0
    
        'Primer caracter
        If InStr(CadenaAux, "`") = 1 Then
             CadenaAux = Right(CadenaAux, Len(CadenaAux) - 1)
        Else
            'Esta al medio
            If InStr(CadenaAux, "`") <> Len(CadenaAux) Then
                CadenaAux = Mid(CadenaAux, 1, InStr(CadenaAux, "`") - 1) & Mid(CadenaAux, InStr(CadenaAux, "`") + 1, Len(CadenaAux))
            Else
            'Es el ultimo
                CadenaAux = Left(CadenaAux, Len(CadenaAux) - 1)
            End If
        End If
        
    Loop
    
    'Analisis de "á"
    Do While InStr(CadenaAux, "á") > 0
    
        'Primer caracter
        If InStr(CadenaAux, "á") = 1 Then
             CadenaAux = Right(CadenaAux, Len(CadenaAux) - 1)
        Else
            'Esta al medio
            If InStr(CadenaAux, "á") <> Len(CadenaAux) Then
                CadenaAux = Mid(CadenaAux, 1, InStr(CadenaAux, "á") - 1) & Mid(CadenaAux, InStr(CadenaAux, "á") + 1, Len(CadenaAux))
            Else
            'Es el ultimo
                CadenaAux = Left(CadenaAux, Len(CadenaAux) - 1)
            End If
        End If
        
    Loop
    
    'Analisis de "é"
    Do While InStr(CadenaAux, "é") > 0
    
        'Primer caracter
        If InStr(CadenaAux, "é") = 1 Then
             CadenaAux = Right(CadenaAux, Len(CadenaAux) - 1)
        Else
            'Esta al medio
            If InStr(CadenaAux, "é") <> Len(CadenaAux) Then
                CadenaAux = Mid(CadenaAux, 1, InStr(CadenaAux, "é") - 1) & Mid(CadenaAux, InStr(CadenaAux, "é") + 1, Len(CadenaAux))
            Else
            'Es el ultimo
                CadenaAux = Left(CadenaAux, Len(CadenaAux) - 1)
            End If
        End If
        
    Loop
    
    'Analisis de "í"
    Do While InStr(CadenaAux, "í") > 0
    
        'Primer caracter
        If InStr(CadenaAux, "í") = 1 Then
             CadenaAux = Right(CadenaAux, Len(CadenaAux) - 1)
        Else
            'Esta al medio
            If InStr(CadenaAux, "í") <> Len(CadenaAux) Then
                CadenaAux = Mid(CadenaAux, 1, InStr(CadenaAux, "í") - 1) & Mid(CadenaAux, InStr(CadenaAux, "í") + 1, Len(CadenaAux))
            Else
            'Es el ultimo
                CadenaAux = Left(CadenaAux, Len(CadenaAux) - 1)
            End If
        End If
        
    Loop
    
    'Analisis de "ó"
    Do While InStr(CadenaAux, "ó") > 0
    
        'Primer caracter
        If InStr(CadenaAux, "ó") = 1 Then
             CadenaAux = Right(CadenaAux, Len(CadenaAux) - 1)
        Else
            'Esta al medio
            If InStr(CadenaAux, "ó") <> Len(CadenaAux) Then
                CadenaAux = Mid(CadenaAux, 1, InStr(CadenaAux, "ó") - 1) & Mid(CadenaAux, InStr(CadenaAux, "ó") + 1, Len(CadenaAux))
            Else
            'Es el ultimo
                CadenaAux = Left(CadenaAux, Len(CadenaAux) - 1)
            End If
        End If
        
    Loop
    
    'Analisis de "ú"
    Do While InStr(CadenaAux, "ú") > 0
    
        'Primer caracter
        If InStr(CadenaAux, "ú") = 1 Then
             CadenaAux = Right(CadenaAux, Len(CadenaAux) - 1)
        Else
            'Esta al medio
            If InStr(CadenaAux, "ú") <> Len(CadenaAux) Then
                CadenaAux = Mid(CadenaAux, 1, InStr(CadenaAux, "ú") - 1) & Mid(CadenaAux, InStr(CadenaAux, "ú") + 1, Len(CadenaAux))
            Else
            'Es el ultimo
                CadenaAux = Left(CadenaAux, Len(CadenaAux) - 1)
            End If
        End If
        
    Loop
    
    QuitarCaracteresErroneos = CadenaAux

End Function

Public Function EscribirRegistro(ByVal fecAsi As String, ByVal D_H As Integer, ByVal Monto As Double, ByVal Linadesc As String, ByVal Cuenta As String, ByVal Desc As String) As String
Dim montoDeta As Double
Dim montoDetaStr As String
Dim detaDesc As String
Dim Linea As String

    'Location - Fijo 408
    Linea = "408"
    'Accouting Period
    Linea = Linea & IIf(EsNulo(fecAsi), "0000", Format(fecAsi, "yymm"))
    'Journal Name - Fijo 20 ""
    Linea = Linea & Format_Str("408-FO-01", 20, True, " ")
    'Record Type - Fijo 2 ""
    Linea = Linea & "DR"
    'Budget Indicator - Fijo 1 ""
    Linea = Linea & "A"
    'Budget Version - Fijo 2 ""
    Linea = Linea & Format_Str("", 2, True, " ")
    'UEC - Fijo 3 ""
    Linea = Linea & Format_Str("", 3, True, " ")
    'CCC - Fijo 3 ""
    Linea = Linea & Format_Str("", 3, True, " ")
    'Other - Fijo 3 ""
    Linea = Linea & Format_Str("", 3, True, " ")
    'Major Account - Fijo 4 ""
    Linea = Linea & Format_Str("", 4, True, " ")
    'Sub Account - Fijo 4 ""
    Linea = Linea & Format_Str("", 4, True, " ")
    'Activity - Fijo 2 ""
    Linea = Linea & Format_Str("", 2, True, " ")
    'CEC - Fijo 3 ""
    Linea = Linea & Format_Str("", 3, True, " ")
    'Sub CEC - Fijo 2 ""
    Linea = Linea & Format_Str("", 2, True, " ")
    'Geo Code - Fijo 2 ""
    Linea = Linea & Format_Str("", 2, True, " ")
    'US/Non US indicator - Fijo 1 ""
    Linea = Linea & Format_Str("", 1, True, " ")
    'Affiliate Code - Fijo 3 ""
    Linea = Linea & Format_Str("", 3, True, " ")
    'MPC - Fijo 2 ""
    Linea = Linea & Format_Str("", 2, True, " ")
    'Reserved - Fijo 4 ""
    Linea = Linea & Format_Str("", 4, True, " ")
    'Debit Credit Indicator
    Select Case D_H
        Case 0 'Debe
            Linea = Linea & "D"
        Case 1 'Haber
            Linea = Linea & "C"
        Case 2 'Variable
            Linea = Linea & "V"
        Case 3 'Variable Inv
            Linea = Linea & "I"
        Case Else
            Linea = Linea & " "
    End Select
    'Accounted Currency Amount - Fijo 19 ""
    Linea = Linea & Format_Str("", 19, True, " ")
    'Entered Currency Amount - Monto del detalle
    montoDeta = IIf(EsNulo(Monto), 0, Monto)
    montoDetaStr = CStr(Round(montoDeta, 3))
    montoDetaStr = Format(montoDetaStr, "##############0.000")
    Linea = Linea & Format_StrNro(montoDetaStr, 19, True, " ")
    'Statistical Amount - Fijo 22 ""
    Linea = Linea & Format_Str("", 22, True, " ")
    'Descricion
    detaDesc = IIf(EsNulo(Linadesc), "", Left(Linadesc, 5))
    detaDesc = detaDesc & "-" & IIf(EsNulo(Desc), "", Trim(Desc))
    Linea = Linea & Format_Str(QuitarCaracteresErroneos(detaDesc), 240, True, " ")
    'Local Acount - Fijo 20 ""
    Linea = Linea & Format_Str("", 20, True, " ")
    'Detail 1 - Fijo 10 ""
    Linea = Linea & Format_Str("", 10, True, " ")
    'Detail 2 - Fijo 10 ""
    Linea = Linea & Format_Str("", 10, True, " ")
    'Detail 3 - Fijo 15 ""
    Linea = Linea & Format_Str("", 15, True, " ")
    'Detail 4 - Fijo 20 ""
    Linea = Linea & Format_Str("", 20, True, " ")
    'Detail 5 - Fijo 20 ""
    Linea = Linea & Format_Str("", 20, True, " ")
    'Detail 6 - Fijo 20 ""
    Linea = Linea & Format_Str("", 20, True, " ")
    'Detail 7 - Fijo 20 ""
    Linea = Linea & Format_Str("", 20, True, " ")
    'Detail 8 - Fijo 30 ""
    Linea = Linea & Format_Str("", 30, True, " ")
    'Detail 9 - Fijo 30 ""
    Linea = Linea & Format_Str("", 30, True, " ")
    'Detail 10 - Fijo 25 ""
    Linea = Linea & Format_Str("", 25, True, " ")
    'Detail 11 - Fijo 25 ""
    Linea = Linea & Format_Str("", 25, True, " ")
    'Detail 12 - Fijo 25 ""
    Linea = Linea & Format_Str("", 25, True, " ")
    'Detail 13 - Fijo 25 ""
    Linea = Linea & Format_Str("", 25, True, " ")
    'Detail 14 - Fijo 25 ""
    Linea = Linea & Format_Str("", 25, True, " ")
    'Detail 15 - Fijo 25 ""
    Linea = Linea & Format_Str("", 25, True, " ")
    'Detail 16 - Fijo 25 ""
    Linea = Linea & Format_Str("", 25, True, " ")
    'Trasaction Qualifier - Fijo 18 ""
    Linea = Linea & Format_Str("", 18, True, " ")
    'Trasaction Qualifier Type - Fijo 1 ""
    Linea = Linea & Format_Str("", 1, True, " ")
    'Original Source Name - Fijo FOL
    Linea = Linea & Format_Str("FOL", 25, True, " ")
    'Currency - Fijo ARS
    Linea = Linea & Format_Str("ARS", 15, True, " ")
    'Conversion Type - Fijo 30 ""
    Linea = Linea & Format_Str("", 30, True, " ")
    'Filler - Fijo 6 ""
    Linea = Linea & Format_Str("", 6, True, " ")
    'Conversion Rate - Fijo 25 ""
    Linea = Linea & Format_Str("", 25, True, " ")
    'Category Name - Fijo 25 ""
    Linea = Linea & Format_Str("", 25, True, " ")
    'Document Sequence Number - Fijo 15 ""
    Linea = Linea & Format_Str("", 15, True, " ")
    'Conversion Date - Fijo 8 ""
    Linea = Linea & Format_Str("", 8, True, " ")
    'LBC - Fijo 04086
    Linea = Linea & "04086"
    'Dept - 5 caracteres de la cuenta a partir de la posicion 10
    Linea = Linea & IIf(EsNulo(Cuenta), Format_Str("", 5, True, " "), Format_Str(Mid(Cuenta, 10, 5), 5, True, " "))
    'Major/Sub Account - 8 caracteres de la cuenta a partir de la posicion 15
    Linea = Linea & IIf(EsNulo(Cuenta), Format_Str("", 8, True, " "), Format_Str(Mid(Cuenta, 15, 8), 8, True, " "))
    'CEC/SE - 4 caracteres de la cuenta a partir de la posicion 23
    Linea = Linea & IIf(EsNulo(Cuenta), Format_Str("", 4, True, " "), Format_Str(Mid(Cuenta, 23, 4), 4, True, " "))
    'LB Affiliate Code - Fijo 00000
    Linea = Linea & "00000"
    'Fiscal Code - Fijo 00001
    Linea = Linea & "00001"
    'Conversion Code - Fijo 0000
    Linea = Linea & "0000"
    'Filler - Fijo 57 ""
    Linea = Linea & Format_Str("", 57, True, " ")
    
    EscribirRegistro = Linea

End Function
