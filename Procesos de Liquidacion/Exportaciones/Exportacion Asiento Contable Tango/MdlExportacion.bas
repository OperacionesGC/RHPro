Attribute VB_Name = "MdlExportacion"
Option Explicit

Global Const Version = "1.01"
Global Const FechaVersion = "21/10/2009"
Global Const UltimaModificacion = "Encriptacion de string connection"
Global Const UltimaModificacion1 = "Manuel Lopez"

Private Type TR_Datos_Varios
    Convenio_Lecop As String        'String   long 4  -
    Filler As String                'String   long 1  -
    Cliente_Ya_Existente As String  'String   long 1  -
End Type

Global IdUser               As String
Global Fecha                As Date
Global Hora                 As String

'Adrián - Declaración de dos nuevos registros.
Global rs_Empresa           As New ADODB.Recordset
Global rs_tipocod           As New ADODB.Recordset

Global Fecha_Inicio_periodo As Date
Global Fecha_Fin_Periodo    As Date
Global StrSql2              As String
Global SeparadorDecimales   As String
Global totalImporte         As Double
Global Total                As Double
Global TotalABS             As Double
Global UltimaLeyenda        As String
Global EsUltimoItem         As Boolean
Global EsUltimoProceso      As Boolean


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
    
    Nombre_Arch = PathFLog & "Exp_Asiento_Contable" & "-" & NroProcesoBatch & ".log"
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Nombre_Arch, True)
    
    ' Obtengo el Process ID
    PID = GetCurrentProcessId
    Flog.writeline "-------------------------------------------------"
    Flog.writeline "Version                  : " & Version
    Flog.writeline "Fecha Ultima Modificacion: " & FechaVersion
    Flog.writeline "Ultima Modificacion      : " & UltimaModificacion
    Flog.writeline "PID                      : " & PID
    Flog.writeline "-------------------------------------------------"
    Flog.writeline
    
    'Cambio el estado del proceso a Procesando y el PID
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcprogreso = 0, bprcestado = 'Procesando', bprcpid = " & PID & " WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE btprcnro = 104 AND bpronro =" & NroProcesoBatch
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

Public Sub Generacion(ByVal bpronro As Long, ByVal Nroliq As Long, ByVal Asinro As String, ByVal Empresa As Long, ByVal ProcVol As Long)
' --------------------------------------------------------------------------------------------
' Descripcion: Procedimiento de generacion del archivo de Asiento Contable
' Autor      : FGZ
' Fecha      : 25/10/2004
' Modificado : 09/08/2005 - Fapitalle N. - Se agrega la generacion de encabezado
'                                        - Se agregan casos en Programa: TAB y CUENTAZ
'                                        - Se agregan casos especiales en Programa para Shering
'                                        - Se agregan casos especiales en Programa para Halliburton
' --------------------------------------------------------------------------------------------
Const ForReading = 1
Const TristateFalse = 0

Dim fExport
Dim fAuxiliar
Dim Directorio As String
Dim ArchivoEnc As String
Dim fArchEnc
Dim ArchivoDet As String
Dim fArchDet
Dim ArchivoUni As String
Dim fArchUni
Dim Intentos As Integer
Dim carpeta

Dim asi_aux3
Dim cod
Dim codb
Dim fecha_asto As String
Dim comentario As String
Dim debe_haber As String
Dim detalle As String
Dim titulo As String
Dim cab As String
Dim cuenta_asto
Dim asto_cuenta
Dim asto_ccos
Dim asto_legajo
Dim orden0(20) As Integer
Dim orden1(20) As Integer
Dim orden2(20) As Integer
Dim I As Integer

Dim strLinea As String
Dim Aux_Linea As String
Dim cadena As String
Dim Aux_str As String
Dim tipo As String
Dim Cantidad As String
Dim Posicion As String
Dim Formato  As String
Dim Nro As Long
Dim NroL As Long
Dim Programa As String
Dim POS As Integer
Dim debeCod As String
Dim haberCod As String
Dim tmpStr As String
Dim separadorCampos
Dim Completa As Boolean
Dim Enter As String
Dim Fecha_Proc As Date

'Registros
Dim rs_Modelo As New ADODB.Recordset
Dim rs_Periodo As New ADODB.Recordset
Dim rs_Procesos As New ADODB.Recordset
Dim rs_Empleado As New ADODB.Recordset

'Archivo de exportacion
StrSql = "SELECT sis_dirsalidas FROM sistema WHERE sisnro = 1 "
OpenRecordset StrSql, rs
If Not rs.EOF Then
    Directorio = Trim(rs!sis_dirsalidas)
End If

StrSql = "SELECT * FROM modelo WHERE modnro = 234"
OpenRecordset StrSql, rs_Modelo
If Not rs_Modelo.EOF Then
    If Not IsNull(rs_Modelo!modarchdefault) Then
        Directorio = Directorio & Trim(rs_Modelo!modarchdefault)
    Else
        Flog.writeline Espacios(Tabulador * 1) & "El modelo no tiene configurada la carpeta desteino. El archivo será generado en el directorio default"
    End If
    SeparadorDecimales = rs_Modelo!modsepdec
    separadorCampos = rs_Modelo!modseparador
Else
    Flog.writeline Espacios(Tabulador * 1) & "No se encontró el modelo. El archivo será generado en el directorio default"
End If

'cargo el periodo
StrSql = "SELECT * FROM periodo WHERE pliqnro = " & CStr(Nroliq)
OpenRecordset StrSql, rs_Periodo
If rs_Periodo.EOF Then
    Flog.writeline Espacios(Tabulador * 1) & "No se encontró el Periodo"
    Exit Sub
End If

  
'Set fs = CreateObject("Scripting.FileSystemObject")
'Activo el manejador de errores
'On Error Resume Next
'Set fExport = fs.CreateTextFile(Archivo, True)
'If Err.Number <> 0 Then
'    Flog.Writeline Espacios(Tabulador * 1) & "La carpeta Destino no existe. Se creará."
'    Set carpeta = fs.CreateFolder(Directorio)
'    Set fExport = fs.CreateTextFile(Archivo, True)
'End If
'desactivo el manejador de errores
'On Error GoTo 0


' Comienzo la transaccion
MyBeginTrans

'Busco los procesos a evaluar
StrSql = "SELECT * FROM  proc_vol"
StrSql = StrSql & " INNER JOIN linea_asi ON proc_vol.vol_cod = linea_asi.vol_cod"
StrSql = StrSql & " INNER JOIN mod_asiento_aux ON linea_asi.masinro = mod_asiento_aux.masinro"
StrSql = StrSql & " WHERE proc_vol.pliqnro =" & Nroliq
If ProcVol <> 0 Then 'si no son todos
    StrSql = StrSql & " AND linea_asi.vol_cod IN (" & ProcVol & ")"
End If
StrSql = StrSql & " AND linea_asi.masinro IN (" & Asinro & ")"
StrSql = StrSql & " AND linea_asi.cuenta <> '999999.999'"
StrSql = StrSql & " ORDER BY mod_asiento_aux.asi_aux3"
OpenRecordset StrSql, rs_Procesos

'seteo de las variables de progreso
Progreso = 0
CConceptosAProc = rs_Procesos.RecordCount
If CConceptosAProc = 0 Then
    CConceptosAProc = 1
    Flog.writeline Espacios(Tabulador * 1) & " No hay Proceso de Volcados para ese asiento en ese periodo"
Else
    Flog.writeline Espacios(Tabulador * 1) & " Lineas de Procesos de Volcados para ese asiento en ese periodo " & CConceptosAProc
End If
IncPorc = (100 / CConceptosAProc)

'Procesamiento
If rs_Procesos.EOF Then
    Flog.writeline Espacios(Tabulador * 2) & "No hay nada que procesar"
End If

'------------------------------------------------------------------------
' Genero el detalle de la exportacion
'------------------------------------------------------------------------

EsUltimoItem = False
EsUltimoProceso = False
asi_aux3 = -1

For I = 0 To 20
orden0(I) = 1
orden1(I) = 1
orden2(I) = 1
Next

cod = 1
codb = 1

Do While Not rs_Procesos.EOF
    If EsUltimoRegistro(rs_Procesos) Then
        EsUltimoProceso = True
    End If
    Nro = Nro + 1 'Contador de Lineas
    Flog.writeline Espacios(Tabulador * 1) & Nro & " -------------------------------------"
    Flog.writeline Espacios(Tabulador * 1) & "Exportando datos del proceso de volcado " & rs_Procesos!vol_cod & " Linea " & rs_Procesos!masinro & " cuenta: " & rs_Procesos!Cuenta
    Cantidad_Warnings = 0

'-*-*-*-*-*-*------------------------------*-*-*-*-*-*-*-*-*---------------
    fecha_asto = Format(rs_Procesos!vol_fec_asiento, "YYYYMMDD")
    
    If rs_Procesos!asi_aux3 <> asi_aux3 Then 'IF FIRST-OF(mod_asiento.asi_aux3)
        asi_aux3 = rs_Procesos!asi_aux3
        If rs_Procesos!masinro = 4 Or rs_Procesos!masinro = 5 Or rs_Procesos!masinro = 6 Then
            ArchivoEnc = "\ASENCP" & fecha_asto & ".txt"
            ArchivoDet = "\ASDETP" & fecha_asto & ".txt"
            cod = 1
            ArchivoUni = "\ASUNIDP" & fecha_asto & ".txt"
            codb = 2
        Else
            ArchivoEnc = "\ASENC" & fecha_asto & ".txt"
            ArchivoDet = "\ASDET" & fecha_asto & ".txt"
            cod = 2
            ArchivoUni = "\ASUNID" & fecha_asto & ".txt"
            codb = 1
        End If
        If rs_Procesos!asi_aux3 < 3 Then
            If rs_Procesos!asi_aux3 = 1 Then
                titulo = rs_Procesos!desclinea
            Else
                titulo = "PREV.SAC - VACACIONES - SAC TC"
            End If
            cab = titulo & String(39 - Len(titulo), " ") & "1.0000 "
            cab = cab & fecha_asto & fecha_asto & String(23, " ") & "8888C"
            'abrir archivo encabezado aca <---
            Call abrir_archivo(Directorio, ArchivoEnc, fArchEnc)
            'escribir cab en archivo encabezado
            fArchEnc.writeline cab
            'cerrar archivo encabezado aca <---
            fArchEnc.Close
        End If
        'abrir archivo detalle aca <---
        Call abrir_archivo(Directorio, ArchivoDet, fArchDet)
        'abrir archivo unid aca <---
        Call abrir_archivo(Directorio, ArchivoUni, fArchUni)
    End If
    
    comentario = rs_Procesos!desclinea 'comentario
    If rs_Procesos!dh Then debe_haber = "D" Else debe_haber = "H" 'debe_haber
    If Mid(rs_Procesos!Cuenta, Len(rs_Procesos!Cuenta) - 1, 1) = "-" Then 'cuenta_asto
        cuenta_asto = Mid(rs_Procesos!Cuenta, 1, Len(rs_Procesos!Cuenta) - 2)
    Else
        cuenta_asto = rs_Procesos!Cuenta
    End If
    If Len(cuenta_asto) = 17 Or Len(cuenta_asto) = 11 Then 'asto_cuenta & asto_ccos
        asto_cuenta = Mid(cuenta_asto, 6, 6)
        asto_ccos = Mid(cuenta_asto, 1, 5)
    Else
        asto_cuenta = cuenta_asto
        asto_ccos = ""
    End If
    If Len(cuenta_asto) = 17 Then asto_legajo = Mid(cuenta_asto, 12, 6) Else asto_legajo = "" ' asto_legajo
    If asto_legajo <> "" Then
        StrSql = "SELECT terape,ternom FROM empleado WHERE empleg = " & asto_legajo 'comentario
        OpenRecordset StrSql, rs_Empleado
        If Not rs_Empleado.EOF Then
            comentario = Trim(comentario) & " - " & rs_Empleado!terape & " - " & rs_Empleado!ternom
        End If
    End If
    
    'escribir la linea de detalle en el archivo detalle
    detalle = item(asto_ccos, 5) & item(asto_cuenta, 6) & item(comentario, 30) & debe_haber & String(20, " ")
    detalle = detalle & Format(rs_Procesos!Monto, "00000000000.00") & "  8888" & Format(orden2(cod), "0000")
    detalle = detalle & Format(rs_Procesos!Monto, "00000000000.00")
    orden2(cod) = orden2(cod) + 1
    'escribir detalle en archivo detalle
    fArchDet.writeline detalle
        
    'escribir la linea de detalle en el archivo unid
    detalle = item(asto_ccos, 5) & item(asto_cuenta, 6) & "0000000001.0000" & debe_haber & String(20, " ")
    detalle = detalle & "  8888" & Format(orden0(codb), "0000")
    detalle = detalle & Format(rs_Procesos!Monto, "000000000.00")
    orden0(codb) = orden0(codb) + 1
    'escribir detalle en archivo unid
    fArchUni.writeline detalle
    
    'si es el ultimo cerrar el archivo detalle y el unid
        rs_Procesos.MoveNext
        If Not rs_Procesos.EOF Then
            If rs_Procesos!asi_aux3 <> asi_aux3 Then
                fArchDet.Close
                fArchUni.Close
            End If
        End If
        rs_Procesos.MovePrevious
        
'-*-*-*-*-*-*------------------------------*-*-*-*-*-*-*-*-*---------------
    'Actualizo el progreso del Proceso
    Progreso = Progreso + IncPorc
    TiempoAcumulado = GetTickCount
    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & _
                 ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & _
                 "' WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    Flog.writeline Espacios(Tabulador * 1) & "--------------------progreso-- " & FormatNumber(Progreso, 2, vbTrue) & " %"
                
    rs_Procesos.MoveNext
Loop

'Fin de la transaccion
MyCommitTrans


If rs_Procesos.State = adStateOpen Then rs_Procesos.Close
If rs_Periodo.State = adStateOpen Then rs_Periodo.Close
If rs_Modelo.State = adStateOpen Then rs_Modelo.Close
If rs_Empleado.State = adStateOpen Then rs_Empleado.Close
Set rs_Procesos = Nothing
Set rs_Periodo = Nothing
Set rs_Modelo = Nothing
Set rs_Empleado = Nothing

Exit Sub
CE:
    Flog.writeline Espacios(Tabulador * 1) & "Error: " & Err.Description
    HuboError = True
    MyRollbackTrans

    If rs_Procesos.State = adStateOpen Then rs_Procesos.Close
    If rs_Periodo.State = adStateOpen Then rs_Periodo.Close
    If rs_Modelo.State = adStateOpen Then rs_Modelo.Close
    Set rs_Procesos = Nothing
    Set rs_Periodo = Nothing
    Set rs_Modelo = Nothing
End Sub

Public Sub LevantarParamteros(ByVal bpronro As Long, ByVal parametros As String)
' --------------------------------------------------------------------------------------------
' Descripcion: Procedimiento para levantar los parametros pasados en batch_proceso en bprcparam
' Autor      : FGZ
' Fecha      :
' Ult. Mod   : 19/8/2005 - Fapitalle N. - Adecuado al proceso de exportacion formato TANGO
' Fecha      :
' --------------------------------------------------------------------------------------------
Dim pos1 As Integer
Dim pos2 As Integer
Dim Separador As String

Dim Periodo As Long
Dim Asiento As String
Dim Empresa As Long
Dim TipoArchivo As Long
Dim ProcVol As Long

'Orden de los parametros
'pliqnro
'Asinro, lista separada por comas
'proceso de volcado, 0=todos

Separador = "@"
' Levanto cada parametro por separado
If Not IsNull(parametros) Then
    If Len(parametros) >= 1 Then
        pos1 = 1
        pos2 = InStr(pos1, parametros, Separador) - 1
        Periodo = Mid(parametros, pos1, pos2 - pos1 + 1)
        
        pos1 = pos2 + 2
        pos2 = InStr(pos1, parametros, Separador) - 1
        Asiento = Mid(parametros, pos1, pos2 - pos1 + 1)
        
        pos1 = pos2 + 2
        pos2 = Len(parametros)
        ProcVol = CLng(Mid(parametros, pos1, pos2 - pos1 + 1))
        
    End If
End If
Call Generacion(bpronro, Periodo, Asiento, Empresa, ProcVol)
End Sub

Public Function item(ByVal cadena As String, ByVal Longitud As Integer)
Dim cad
    cad = Left(cadena, Longitud)
    If Len(cad) < Longitud Then
        cad = cad & String(Longitud - Len(cad), " ")
    End If
    item = cad
End Function

Public Sub abrir_archivo(ByVal dir As String, ByVal nom As String, ByRef archivo)
Const ForAppending = 8
Dim fs
Dim faux
Dim carpeta

Set fs = CreateObject("Scripting.FileSystemObject")
'Activo el manejador de errores
On Error Resume Next
Flog.writeline Espacios(Tabulador * 1) & "Carpeta de Trabajo: " & dir
Set archivo = fs.OpenTextFile(dir & nom, ForAppending)
If Err.Number <> 0 Then
    Flog.writeline Espacios(Tabulador * 1) & "Se creará el archivo " & nom
    Set carpeta = fs.CreateFolder(dir)
    Set archivo = fs.CreateTextFile(dir & nom, True)
Else
    Flog.writeline Espacios(Tabulador * 1) & "Se añadirá información al archivo " & nom
End If
Flog.writeline Espacios(Tabulador * 1)
'desactivo el manejador de errores
On Error GoTo 0
End Sub

Public Sub Archivo_ASTO_SAP(ByVal dir As String, ByVal Fecha As Date)
'--------------------------------------------------------------------------------
'  Archivo:
'  Descripci¢n: Genera el archivo ASTOmmaa.txt para el volcado SAP de Halliburton
'  Autor: Fapitalle N.
'  Fecha: 18/08/2005
'-------------------------------------------------------------------------------
Dim fAstoSAP
Dim fs
Dim archivo
Dim carpeta
Dim cadena As String

archivo = dir & "\ASTO" & Format(Fecha, "MMYY") & ".txt"
Set fs = CreateObject("Scripting.FileSystemObject")
'Activo el manejador de errores
On Error Resume Next
Set fAstoSAP = fs.CreateTextFile(archivo, True)
If Err.Number <> 0 Then
    Set carpeta = fs.CreateFolder(dir)
    Set fAstoSAP = fs.CreateTextFile(archivo, True)
End If
'desactivo el manejador de errores
On Error GoTo 0

cadena = "Constante" + ";" + "Blancos" + ";" + "Cuenta" + ";" + "Blancos" + ";" + "Entidad" + ";" + _
         "Blancos" + ";" + "Constante" + ";" + "Blanco" + ";" + "Moneda" + ";" + "Blanco" + ";" + _
         "Constante" + ";" + "Descripcion" + ";" + "Blanco" + ";" + "Fecha" + ";" + "Blanco" + ";" + _
         "Constante" + ";" + "Blanco" + ";" + "Importe" + ";" + "Debe/Haber"

fAstoSAP.writeline cadena
fAstoSAP.Close
    
End Sub

Private Sub Cuenta(ByVal Cuenta As String, ByRef Str_Salida As String)
'--------------------------------------------------------------------------------
'  Archivo: conta/cuenta.p
'  Descripci¢n: devuelve la cuenta de la linea del asiento de la siguiente manera:
'               999999999999.999999.99999999
'               ej: si la cuenta es: 11000003.521110.01
'                   debera salir:000011000003.521110.00000001
'  Autor: FGZ
'  Fecha: 26/10/2004
'-------------------------------------------------------------------------------
Dim I As Integer
Dim cadena As String

    cadena = ""
    I = 1
    Do While I <= Len(Cuenta)
        cadena = cadena + Mid(Cuenta, I, 1)
        I = I + 1
    Loop
    cadena = cadena & IIf(Len(cadena) = 10, ".1000", "")
    Str_Salida = cadena

End Sub


Private Sub Cuenta_1(ByVal Cuenta As String, ByRef Str_Salida As String)
'--------------------------------------------------------------------------------
'  Archivo: conta/cuenta.p
'  Descripci¢n: devuelve la cuenta de la linea del asiento de la siguiente manera:
'               999999999999.999999.99999999
'               ej: si la cuenta es: 11000003.521110.01
'                   debera salir:000011000003.521110.00000001
'  OBS  : TRAER LA UNIDAD DE NEGOCIO AL FRENTE DE LA CUENTA
'  Autor: FGZ
'  Fecha: 26/10/2004
'-------------------------------------------------------------------------------
Dim I As Integer
Dim cadena As String
Dim subcad As String

    cadena = ""
    I = 1
    Do While I <= Len(Cuenta)
        cadena = cadena + Mid(Cuenta, I, 1)
        I = I + 1
    Loop
    cadena = cadena & IIf(Len(cadena) = 10, ".1000", "")
    
    'PARA TRAER LA UNIDAD DE NEGOCIO AL FRENTE DE LA CUENTA
    subcad = cadena
    cadena = ""
    cadena = Mid(subcad, 12, 4) & "." & Mid(subcad, 1, 10)
    
    Str_Salida = cadena
End Sub


Private Sub Cuenta_2(ByVal Cuenta As String, ByRef Str_Salida As String)
'--------------------------------------------------------------------------------
'  Archivo: conta/cuenta2.p
'  Descripci¢n: devuelve los primeros numeros de la cuenta, hasta el primer punto
'               en un formato de 12 digitos.
'  Autor: FGZ
'  Fecha: 26/10/2004
'-------------------------------------------------------------------------------
Dim I As Integer
Dim cadena As String

    cadena = ""
    I = 1
    Do While I <= Len(Cuenta)
        cadena = cadena + Mid(Cuenta, I, 1)
        I = I + 1
    Loop
    
    cadena = IIf(Mid(cadena, 12, 4) = "", "1000", Mid(cadena, 12, 4))
    Str_Salida = cadena

End Sub

Private Sub Cuenta_3(ByVal Cuenta As String, ByRef Str_Salida As String)
'--------------------------------------------------------------------------------
'  Archivo: conta/cuenta2.p
'  Descripci¢n: devuelve los primeros numeros de la cuenta, hasta el primer punto
'               en un formato de 6 digitos.
'  Autor: FGZ
'  Fecha: 26/10/2004
'-------------------------------------------------------------------------------
Dim I As Integer
Dim cadena As String

    cadena = ""
    I = 1
    Do While I <= Len(Cuenta)
        cadena = cadena + Mid(Cuenta, I, 1)
        I = I + 1
    Loop
    cadena = IIf(Mid(cadena, 1, 6) = "", "000000", Mid(cadena, 1, 6))
    Str_Salida = cadena

End Sub

Private Sub Cuenta_4(ByVal Cuenta As String, ByRef Str_Salida As String)
'--------------------------------------------------------------------------------
'  Archivo: conta/cuenta.p
'  Descripci¢n: devuelve la cuenta de la linea del asiento de la siguiente manera:
'               999999999999.999999.99999999
'               ej: si la cuenta es: 11000003.521110.01
'                   debera salir:000011000003.521110.00000001
'  Autor: FGZ
'  Fecha: 26/10/2004
'-------------------------------------------------------------------------------
Dim I As Integer
Dim cadena As String

    cadena = ""
    I = 1
    Do While I <= Len(Cuenta)
        cadena = cadena + Mid(Cuenta, I, 1)
        I = I + 1
    Loop
    cadena = IIf(Mid(cadena, 8, 3) = "", "000", Mid(cadena, 8, 3))
    Str_Salida = cadena

End Sub


Private Sub Importe(ByVal Monto As Double, ByVal Debe As Boolean, ByVal Completar As Boolean, ByVal Longitud As Integer, ByRef Str_Salida As String)
'--------------------------------------------------------------------------------
'  Archivo: conta/importe.p
'  Descripci¢n: devuelve el importe de la linea en el siguiente formato:
'               Si va al debe es + y - sino, con dos decimales seguidos sin
'               Separador
'  Autor: FGZ
'  Fecha: 26/10/2004
'-------------------------------------------------------------------------------
Dim I As Integer
Dim cadena As String
Dim Aux_Cadena As String
Dim Parte_Entera As String
Dim Parte_Decimal As String
Dim Numero
Dim Balancea As Boolean
Dim Diferencia As Single

Balancea = False
Do While Not Balancea
    Numero = Split(CStr(Monto), ".")
    Parte_Entera = Fix(Monto)
    'Parte_Decimal = CStr(Format(IIf(Round((Monto - Parte_Entera) * 100, 0) > 0, Round((Monto - Parte_Entera) * 100, 0), 0), "##"))
    Parte_Decimal = CStr(Format(IIf(Round(Abs((Monto - Parte_Entera)) * 100, 0) <> 0, Round(Abs(Monto - Parte_Entera) * 100, 0), 0), "##"))
    If Len(Parte_Decimal) < 2 Then
        Parte_Decimal = "0" & Parte_Decimal
    Else
        Parte_Decimal = Left(Parte_Decimal, 2)
    End If
    
    Numero(0) = Parte_Entera
    If Debe Then
        If Completar Then
            cadena = Format(Numero(0), String(Longitud - 3, "0"))
        Else
            cadena = Numero(0)
        End If
    Else
        If Completar Then
            cadena = Format(Numero(0), String(Longitud - 3, "0"))
        Else
            cadena = Numero(0)
        End If
    End If
    If UBound(Numero) > 0 Then
        Numero(1) = Parte_Decimal
        cadena = cadena & SeparadorDecimales & Left(Numero(1) & "00", 2)
    Else
        cadena = cadena & SeparadorDecimales & "00"
    End If
    
    'Para calcular el total
    If Debe Then
        Aux_Cadena = Numero(0) & "."
    Else
        Aux_Cadena = Numero(0) & "."
    End If
    If UBound(Numero) > 0 Then
        Aux_Cadena = Aux_Cadena & Left(Numero(1) & "00", 2)
    Else
        Aux_Cadena = Aux_Cadena & "00"
    End If
    Aux_Cadena = Replace(Aux_Cadena, ",", ".")
    totalImporte = CDbl(Round(CDbl(totalImporte) + Abs(CDbl(cadena)), 2))
    
    'FGZ - 17/06/2005
    If EsUltimoItem And EsUltimoProceso Then
        Diferencia = Round(Total + CDbl(Aux_Cadena), 2)
        If Diferencia <> 0 Then
            totalImporte = Round(totalImporte - Abs(CDbl(cadena)), 2)
            'Monto = CSng(Aux_Cadena) + Diferencia
            Monto = -1 * Total
        Else
'            Total = Total + CSng(cadena)
'            Balancea = True
            If Debe Then
                Total = Round(Total + CDbl(Abs(Aux_Cadena)), 2)
            Else
                Total = Round(Total - CDbl(Abs(Aux_Cadena)), 2)
            End If
            Balancea = True
        End If
    Else
        If Debe Then
            Total = Round(Total + CDbl(Abs(Aux_Cadena)), 2)
        Else
            Total = Round(Total - CDbl(Abs(Aux_Cadena)), 2)
        End If
        Balancea = True
'
'        Balancea = True
'        Total = Total + CSng(cadena)
    End If
Loop

'cadena = Aux_Cadena
If Completar Then
    If Len(cadena) < Longitud Then
        cadena = String(Longitud - Len(cadena), "0") & cadena
    Else
        If Len(cadena) > Longitud Then
            cadena = Right(cadena, Longitud)
        End If
    End If
End If
Str_Salida = cadena
 
End Sub

Private Sub Fecha1(ByVal Fecha As Date, ByRef Str_Salida As String)
'--------------------------------------------------------------------------------
'  Archivo: conta/importe.p
'  Descripci¢n: devuelve la fecha de emision  en el siguiente formato:
'               999999 donde los primeros tres corresponden al a¤o (099 para
'               1999 y 100 para 2000) y los otros tres digitos son para
'               los dias del año (del 001 al 365).
'  Autor: FGZ
'  Fecha: 26/10/2004
'-------------------------------------------------------------------------------
Dim I As Integer
Dim cadena As String

    cadena = Format(Year(CDate(Fecha)) Mod 1900, "000") & Format(CDate(Fecha) - CDate("01/01/" & Year(CDate(Fecha))), "000")
        
    Str_Salida = cadena

End Sub

Private Sub Fecha2(ByVal Fecha As Date, ByRef Str_Salida As String)
'--------------------------------------------------------------------------------
'  Archivo: conta/importe.p
'  Descripci¢n: devuelve la fecha de emision  en el siguiente formato:
'               999999 donde los primeros tres corresponden al a¤o (099 para
'               1999 y 100 para 2000) y los otros tres digitos son para
'               los dias del año (del 001 al 365).
'  Autor: FGZ
'  Fecha: 26/10/2004
'-------------------------------------------------------------------------------
Dim I As Integer
Dim cadena As String

    cadena = Format(Fecha, "ddmmyy")
        
    Str_Salida = cadena

End Sub

Private Sub Fecha3(ByVal Fecha As Date, ByRef Str_Salida As String)
'--------------------------------------------------------------------------------
'  Archivo: conta/importe.p
'  Descripci¢n: devuelve la fecha de emision  en el siguiente formato:
'               999999 donde los primeros tres corresponden al a¤o (099 para
'               1999 y 100 para 2000) y los otros tres digitos son para
'               los dias del año (del 001 al 365).
'  Autor: FGZ
'  Fecha: 26/10/2004
'-------------------------------------------------------------------------------
Dim I As Integer
Dim cadena As String

    cadena = Format(Fecha, "ddmm")
        
    Str_Salida = cadena

End Sub

Private Sub Fecha4(ByVal Fecha As Date, ByRef Str_Salida As String)
'--------------------------------------------------------------------------------
'  Archivo: conta/importe.p
'  Descripci¢n: devuelve la fecha de emision en el siguiente formato:
'               MYYYY donde los primeros 2 corresponden al mes de 1 a 12
'               y los otros 4 digitos son para el año
'  Autor: Fapitalle N.
'  Fecha: 09/08/2005
'-------------------------------------------------------------------------------
Dim I As Integer
Dim cadena As String

    cadena = Format(Fecha, "MYYYY")
        
    Str_Salida = cadena

End Sub


Private Sub Fecha_Estandar(ByVal Fecha As Date, ByVal Formato As String, ByVal Completar As Boolean, ByVal Longitud As Integer, ByRef Str_Salida As String)
'--------------------------------------------------------------------------------
'  Archivo: conta/importe.p
'  Descripci¢n: devuelve la fecha de emision  en el siguiente formato:
'               999999 donde los primeros tres corresponden al a¤o (099 para
'               1999 y 100 para 2000) y los otros tres digitos son para
'               los dias del año (del 001 al 365).
'  Autor: FGZ
'  Fecha: 26/10/2004
'-------------------------------------------------------------------------------
Dim I As Integer
Dim cadena As String

    cadena = Format(Fecha, Formato)
    'Cadena = Format(Fecha, "ddmmyy")
        
    If Completar Then
        If Len(cadena) < Longitud Then
            cadena = cadena & String(Longitud - Len(cadena), " ")
        End If
    End If
        
    Str_Salida = cadena

End Sub

Private Sub NroLinea(ByVal Linea As Long, ByVal Completar As Boolean, ByVal Longitud As Integer, ByRef Str_Salida As String)
'--------------------------------------------------------------------------------
'  Archivo: conta/nrolinea.p
'  Descripci¢n: .
'  Autor: FGZ
'  Fecha: 26/10/2004
'-------------------------------------------------------------------------------
Dim I As Integer
Dim cadena As String

    cadena = Format(Linea, String(Longitud, "0"))
        
'    If Completar Then
'        If Len(Cadena) < Longitud Then
'            Cadena = String(Longitud - Len(Cadena), "0") & Cadena
'        End If
'    End If
    Str_Salida = cadena
End Sub

Private Sub NroAsiento(ByVal Asiento As Long, ByVal Completar As Boolean, ByVal Longitud As Integer, ByRef Str_Salida As String)
'--------------------------------------------------------------------------------
'  Descripci¢n: .
'  Autor: FGZ
'  Fecha: 26/10/2004
'-------------------------------------------------------------------------------
Dim I As Integer
Dim cadena As String

    cadena = Format(Asiento, String(Longitud, "0"))
        
    Str_Salida = cadena
End Sub


Private Sub Leyenda(ByVal Descripcion As String, ByVal POS As Integer, ByVal Cant As Integer, ByVal Completar As Boolean, ByVal Longitud As Integer, ByRef Str_Salida As String)
'--------------------------------------------------------------------------------
'  Archivo: conta/leyasiento.p
'  Descripci¢n: devuelve el la descripcion.
'  Autor: FGZ
'  Fecha: 26/10/2004
'-------------------------------------------------------------------------------
Dim cadena As String

    If Len(Descripcion) < Cant Then
        cadena = Mid(Descripcion, POS, Len(Descripcion))
    Else
        cadena = Mid(Descripcion, POS, Cant)
    End If

    If Completar Then
        If Len(cadena) < Longitud Then
            cadena = cadena & String(Longitud - Len(cadena), " ")
        End If
    End If
    Str_Salida = cadena

End Sub



Private Sub Leyenda1(ByVal Asi_Cod As Long, ByVal Linea As Long, ByVal Descripcion As String, ByVal POS As Long, ByVal Cant As Integer, ByVal Completar As Boolean, ByVal Longitud As Integer, ByRef Str_Salida As String)
'--------------------------------------------------------------------------------
'  Archivo: conta/leyasiento.p
'  Descripci¢n: devuelve el la descripcion.
'  Autor: FGZ
'  Fecha: 26/10/2004
'-------------------------------------------------------------------------------
Dim cadena As String
Dim rs_Mod_Linea As New ADODB.Recordset
Dim Encontro As Boolean

    If Len(Descripcion) < Cant Then
        cadena = Mid(Descripcion, POS, Len(Descripcion))
    Else
        cadena = Mid(Descripcion, POS, Cant)
    End If

Encontro = False
StrSql = "SELECT * FROM mod_linea "
StrSql = StrSql & " WHERE mod_linea.masinro =" & Asi_Cod
StrSql = StrSql & " AND mod_linea.linaorden =" & Linea
OpenRecordset StrSql, rs_Mod_Linea
If Not rs_Mod_Linea.EOF Then
    '1er nivel de estructura
    If Not EsNulo(rs_Mod_Linea!lineanivternro1) Then
        If rs_Mod_Linea!lineanivternro1 = 32 Then
            Encontro = True
        End If
    End If
    '2do nivel de estructura
    If Not EsNulo(rs_Mod_Linea!lineanivternro2) Then
        If rs_Mod_Linea!lineanivternro2 = 32 Then
            Encontro = True
        End If
    End If
    '3er nivel de estructura
    If Not EsNulo(rs_Mod_Linea!lineanivternro3) Then
        If rs_Mod_Linea!lineanivternro3 = 32 Then
            Encontro = True
        End If
    End If
    
    If Encontro Then
        cadena = "Jornales " & cadena
    Else
        cadena = "Sueldos " & cadena
    End If
End If
    
If Completar Then
    If Len(cadena) < Longitud Then
        cadena = cadena & String(Longitud - Len(cadena), " ")
    End If
End If
Str_Salida = cadena

End Sub

Private Sub Leyenda2(ByVal Asi_Cod As Long, ByVal Linea As Long, ByVal Descripcion As String, ByVal POS As Long, ByVal Cant As Integer, ByVal Completar As Boolean, ByVal Longitud As Integer, ByRef Str_Salida As String)
'--------------------------------------------------------------------------------
'  Archivo:
'  Descripci¢n: devuelve la descripcion del modelo.
'  Autor: DOS
'  Fecha: 18/05/2005
'-------------------------------------------------------------------------------
Dim cadena As String
Dim rs_Mod_Asiento As New ADODB.Recordset

    StrSql = "SELECT * FROM mod_asiento "
    StrSql = StrSql & " WHERE masinro =" & Asi_Cod
    
    OpenRecordset StrSql, rs_Mod_Asiento
    
    cadena = ""
    
    If Not rs_Mod_Asiento.EOF Then
       cadena = rs_Mod_Asiento!masidesc
    End If
    
    rs_Mod_Asiento.Close
        
    If Completar Then
        If Len(cadena) < Longitud Then
            cadena = cadena & String(Longitud - Len(cadena), " ")
        End If
    End If
    
    Str_Salida = cadena

End Sub


Private Sub Leyenda3(ByVal Asi_Cod As Long, ByVal Linea As Long, ByVal Descripcion As String, ByVal POS As Long, ByVal Cant As Integer, ByVal Completar As Boolean, ByVal Longitud As Integer, ByVal periodoMes As Integer, ByVal periodoAnio As Integer, ByRef Str_Salida As String)
'--------------------------------------------------------------------------------
'  Archivo:
'  Descripci¢n: devuelve la descripcion del modelo y el periodo.
'  Autor: DOS
'  Fecha: 18/05/2005
'-------------------------------------------------------------------------------
Dim cadena As String
Dim rs_Mod_Asiento As New ADODB.Recordset

    StrSql = "SELECT * FROM mod_asiento "
    StrSql = StrSql & " WHERE masinro =" & Asi_Cod
    
    OpenRecordset StrSql, rs_Mod_Asiento
    
    cadena = ""
    
    If Not rs_Mod_Asiento.EOF Then
       cadena = Left(rs_Mod_Asiento!masidesc, 7)
    End If
    
    rs_Mod_Asiento.Close
    
    If periodoMes < 10 Then
       cadena = cadena & " 0" & periodoMes
    Else
       cadena = cadena & " " & periodoMes
    End If
    
    cadena = cadena & " " & periodoAnio
        
    If Completar Then
        If Len(cadena) < Longitud Then
            cadena = cadena & String(Longitud - Len(cadena), " ")
        End If
    End If
    
    Str_Salida = cadena

End Sub


Private Sub NroPeriodo(ByVal Periodo As Long, ByVal Inicial As Long, ByVal Completar As Boolean, ByVal Longitud As Integer, ByRef Str_Salida As String)
'--------------------------------------------------------------------------------
'  Descripci¢n: .
'  Autor: FGZ
'  Fecha: 26/10/2004
'-------------------------------------------------------------------------------
Dim I As Integer
Dim cadena As String

    cadena = Format(Periodo + Inicial, String(Longitud, "0"))
        
    Str_Salida = cadena
End Sub


Private Sub ImporteABS(ByVal Monto As Double, ByVal Debe As Boolean, ByVal Completar As Boolean, ByVal Longitud As Integer, ByRef Str_Salida As String)
'--------------------------------------------------------------------------------
'  Archivo: conta/importe.p
'  Descripci¢n: devuelve el importe de la linea en el siguiente formato:
'               el monto esta expresado en valor absoluto, con dos decimales seguidos y
'               el separador de decimales es el definido en el modelo.
'  Autor: FGZ
'  Fecha: 26/10/2004
'-------------------------------------------------------------------------------
Dim I As Integer
Dim cadena As String
Dim Aux_Cadena As String

Dim Parte_Entera As String
Dim Parte_Decimal As String
Dim Numero
Dim Balancea As Boolean
Dim Diferencia As Single

Balancea = False
Do While Not Balancea
    Numero = Split(CStr(Monto), ".")
    Parte_Entera = Fix(Monto)
    Parte_Decimal = CStr(Format(IIf(Round((Monto - Parte_Entera) * 100, 0) <> 0, Round(Abs(Monto - Parte_Entera) * 100, 0), 0), "##"))
    If Len(Parte_Decimal) < 2 Then
        Parte_Decimal = "0" & Parte_Decimal
    End If
    Numero(0) = Parte_Entera
    
    If Completar Then
        cadena = Format(Numero(0), String(Longitud - 3, "0")) & SeparadorDecimales
    Else
        cadena = Numero(0) & SeparadorDecimales
    End If
    If UBound(Numero) > 0 Then
        Numero(1) = Parte_Decimal
        cadena = cadena & Left(Numero(1) & "00", 2)
    Else
        cadena = cadena & "00"
    End If
    cadena = Replace(cadena, ",", ".")
    cadena = Replace(cadena, "-", "")
        
    'Para calcular el total
    If Debe Then
        Aux_Cadena = Numero(0) & "."
    Else
        Aux_Cadena = Numero(0) & "."
    End If
    If UBound(Numero) > 0 Then
        Aux_Cadena = Aux_Cadena & Numero(1)
    Else
        Aux_Cadena = Aux_Cadena & "00"
    End If
    Aux_Cadena = Replace(Aux_Cadena, ",", ".")
    totalImporte = Round(totalImporte + Abs(CDbl(Aux_Cadena)), 2)
    Total = Round(Total + CDbl(cadena), 2)
    'FGZ - 17/06/2005
    If EsUltimoItem And EsUltimoProceso Then
        Diferencia = Round(TotalABS + CDbl(Aux_Cadena), 2)
        If Diferencia <> 0 Then
            totalImporte = Round(totalImporte - Abs(CDbl(Aux_Cadena)), 2)
            Total = Round(Total - CDbl(cadena), 2)
            If Diferencia < 0 Then
                'Monto = CSng(Aux_Cadena) - Diferencia
                Monto = TotalABS * -1
            Else
                'Monto = CSng(Aux_Cadena) + Diferencia
                Monto = -1 * TotalABS
            End If
        Else
            If Debe Then
                TotalABS = Round(TotalABS + CDbl(Aux_Cadena), 2)
            Else
                TotalABS = Round(TotalABS - CDbl(Aux_Cadena), 2)
            End If
            Balancea = True
        End If
    Else
        Balancea = True
        If Debe Then
            TotalABS = Round(TotalABS + CDbl(Aux_Cadena), 2)
        Else
            TotalABS = Round(TotalABS - CDbl(Aux_Cadena), 2)
        End If
    End If
Loop

    
    If Completar Then
        If Len(cadena) < Longitud Then
            cadena = String(Longitud - Len(cadena), "0") & cadena
        End If
    End If
    Str_Salida = cadena

End Sub


Private Sub ImporteABS_2(ByVal Monto As Double, ByVal Debe As Boolean, ByVal Completar As Boolean, ByVal Longitud As Integer, ByRef Str_Salida As String)
'--------------------------------------------------------------------------------
'  Archivo: conta/importe.p
'  Descripci¢n: devuelve el importe de la linea en el siguiente formato:
'               el monto esta expresado en valor absoluto, con dos decimales seguidos y
'               el separador de decimales es el ".", el relleno es con espacios al final
'               La funcion es una ligera modificacion de ImporteABS
'  Autor: Fapitalle N.
'  Fecha: 10/08/2005
'-------------------------------------------------------------------------------
Dim I As Integer
Dim cadena As String
Dim Aux_Cadena As String

Dim Parte_Entera As String
Dim Parte_Decimal As String
Dim Numero
Dim Balancea As Boolean
Dim Diferencia As Single

Balancea = False
Do While Not Balancea
    Numero = Split(CStr(Monto), ".")
    Parte_Entera = Fix(Monto)
    Parte_Decimal = CStr(Format(IIf(Round((Monto - Parte_Entera) * 100, 0) <> 0, Round(Abs(Monto - Parte_Entera) * 100, 0), 0), "##"))
    If Len(Parte_Decimal) < 2 Then
        Parte_Decimal = "0" & Parte_Decimal
    End If
    Numero(0) = Parte_Entera
    
    cadena = Numero(0) & "."
    If UBound(Numero) > 0 Then
        Numero(1) = Parte_Decimal
        cadena = cadena & Left(Numero(1) & "00", 2)
    Else
        cadena = cadena & "00"
    End If
    cadena = Replace(cadena, ",", ".")
    cadena = Replace(cadena, "-", "")
    
    If Completar Then
        cadena = cadena & String(Longitud - Len(cadena), " ")
    End If
    
    'Para calcular el total
    If Debe Then
        Aux_Cadena = Numero(0) & "."
    Else
        Aux_Cadena = Numero(0) & "."
    End If
    If UBound(Numero) > 0 Then
        Aux_Cadena = Aux_Cadena & Numero(1)
    Else
        Aux_Cadena = Aux_Cadena & "00"
    End If
    Aux_Cadena = Replace(Aux_Cadena, ",", ".")
    totalImporte = Round(totalImporte + Abs(CDbl(Aux_Cadena)), 2)
    Total = Round(Total + CDbl(cadena), 2)
    'FGZ - 17/06/2005
    If EsUltimoItem And EsUltimoProceso Then
        Diferencia = Round(TotalABS + CDbl(Aux_Cadena), 2)
        If Diferencia <> 0 Then
            totalImporte = Round(totalImporte - Abs(CDbl(Aux_Cadena)), 2)
            Total = Round(Total - CDbl(cadena), 2)
            If Diferencia < 0 Then
                'Monto = CSng(Aux_Cadena) - Diferencia
                Monto = TotalABS * -1
            Else
                'Monto = CSng(Aux_Cadena) + Diferencia
                Monto = -1 * TotalABS
            End If
        Else
            If Debe Then
                TotalABS = Round(TotalABS + CDbl(Aux_Cadena), 2)
            Else
                TotalABS = Round(TotalABS - CDbl(Aux_Cadena), 2)
            End If
            Balancea = True
        End If
    Else
        Balancea = True
        If Debe Then
            TotalABS = Round(TotalABS + CDbl(Aux_Cadena), 2)
        Else
            TotalABS = Round(TotalABS - CDbl(Aux_Cadena), 2)
        End If
    End If
Loop

    Str_Salida = cadena

End Sub

Private Sub ImporteABS_3(ByVal Monto As Double, ByVal Debe As Boolean, ByVal Completar As Boolean, ByVal Longitud As Integer, ByRef Str_Salida As String)
'--------------------------------------------------------------------------------
'  Archivo: conta/importe.p
'  Descripci¢n: devuelve el importe de la linea en el siguiente formato:
'               el monto esta expresado en valor absoluto, con dos decimales seguidos y
'               el separador de decimales es el ",", el relleno es con ceros adelante
'               La funcion es una ligera modificacion de ImporteABS
'  Autor: Fapitalle N.
'  Fecha: 10/08/2005
'-------------------------------------------------------------------------------
Dim I As Integer
Dim cadena As String
Dim Aux_Cadena As String

Dim Parte_Entera As String
Dim Parte_Decimal As String
Dim Numero
Dim Balancea As Boolean
Dim Diferencia As Single

Balancea = False
'Do While Not Balancea
    Numero = Split(CStr(Monto), ".")
    Parte_Entera = Fix(Monto)
    Parte_Decimal = CStr(Format(IIf(Round((Monto - Parte_Entera) * 100, 0) <> 0, Round(Abs(Monto - Parte_Entera) * 100, 0), 0), "##"))
    If Len(Parte_Decimal) < 2 Then
        Parte_Decimal = "0" & Parte_Decimal
    End If
    Numero(0) = Parte_Entera
    
    cadena = Numero(0) & ","
    If UBound(Numero) > 0 Then
        Numero(1) = Parte_Decimal
        cadena = cadena & Left(Numero(1) & "00", 2)
    Else
        cadena = cadena & "00"
    End If
    cadena = Replace(cadena, "-", "")
    
    If Completar Then
        cadena = String(Longitud - Len(cadena), "0") + cadena
    End If
    
    'Para calcular el total
'    If Debe Then
'        Aux_Cadena = Numero(0) & "."
'    Else
'        Aux_Cadena = Numero(0) & "."
'    End If
'    If UBound(Numero) > 0 Then
'        Aux_Cadena = Aux_Cadena & Numero(1)
'    Else
'        Aux_Cadena = Aux_Cadena & "00"
'    End If
'    Aux_Cadena = Replace(Aux_Cadena, ",", ".")
'    totalImporte = Round(totalImporte + Abs(CDbl(Aux_Cadena)), 2)
'    Total = Round(Total + CDbl(cadena), 2)
    'FGZ - 17/06/2005
'    If EsUltimoItem And EsUltimoProceso Then
'        Diferencia = Round(TotalABS + CDbl(Aux_Cadena), 2)
'        If Diferencia <> 0 Then
'            totalImporte = Round(totalImporte - Abs(CDbl(Aux_Cadena)), 2)
'            Total = Round(Total - CDbl(cadena), 2)
'            If Diferencia < 0 Then
'                'Monto = CSng(Aux_Cadena) - Diferencia
'                Monto = TotalABS * -1
'            Else
'                'Monto = CSng(Aux_Cadena) + Diferencia
'                Monto = -1 * TotalABS
'            End If
'        Else
'            If Debe Then
'                TotalABS = Round(TotalABS + CDbl(Aux_Cadena), 2)
'            Else
'                TotalABS = Round(TotalABS - CDbl(Aux_Cadena), 2)
'            End If
''            Balancea = True
''        End If
'    Else
'        Balancea = True
'        If Debe Then
'            TotalABS = Round(TotalABS + CDbl(Aux_Cadena), 2)
'        Else
'            TotalABS = Round(TotalABS - CDbl(Aux_Cadena), 2)
'        End If
'    End If
'Loop

    Str_Salida = cadena

End Sub


Private Sub ImporteABS_old(ByVal Monto As Single, ByVal Debe As Boolean, ByVal Completar As Boolean, ByVal Longitud As Integer, ByRef Str_Salida As String)
'--------------------------------------------------------------------------------
'  Archivo: conta/importe.p
'  Descripci¢n: devuelve el importe de la linea en el siguiente formato:
'               el monto esta expresado en valor absoluto, con dos decimales seguidos y
'               el separador de decimales es el definido en el modelo.
'  Autor: FGZ
'  Fecha: 26/10/2004
'-------------------------------------------------------------------------------
Dim I As Integer
Dim cadena As String
Dim Aux_Cadena As String

Dim Parte_Entera
Dim Parte_Decimal
Dim Numero

    Numero = Split(CStr(Monto), ".")
    'Parte_Entera = Fix(Monto)
    'Parte_Decimal = IIf(Round((Monto - Parte_Entera) * 100, 0) > 0, Round((Monto - Parte_Entera) * 100, 0), 0)

'    If Debe Then
'       totalImporte = totalImporte + Monto
'    Else
'       totalImporte = totalImporte - Monto
'    End If
    
    If Completar Then
        cadena = Format(Numero(0), String(Longitud - 3, "0")) & SeparadorDecimales
    Else
        cadena = Numero(0) & SeparadorDecimales
    End If
    If UBound(Numero) > 0 Then
        cadena = cadena & Left(Numero(1) & "00", 2)
    Else
        cadena = cadena & "00"
    End If
    cadena = Replace(cadena, ",", ".")
    cadena = Replace(cadena, "-", "")
        
    'Para calcular el total
    If Debe Then
        Aux_Cadena = Numero(0) & "."
    Else
        Aux_Cadena = Numero(0) & "."
    End If
    If UBound(Numero) > 0 Then
        Aux_Cadena = Aux_Cadena & Left(Numero(1) & "00", 2)
    Else
        Aux_Cadena = Aux_Cadena & "00"
    End If
    Aux_Cadena = Replace(Aux_Cadena, ",", ".")
    totalImporte = totalImporte + Abs(CSng(Aux_Cadena))
    'totalImporte = totalImporte + Abs(CSng(cadena))
    
    If Completar Then
        If Len(cadena) < Longitud Then
            cadena = String(Longitud - Len(cadena), "0") & cadena
        End If
    End If
    Str_Salida = cadena

End Sub

Private Sub DebeHaber(ByVal Debe As Boolean, ByVal debeCod As String, ByVal haberCod As String, ByVal Completar As Boolean, ByVal Longitud As Integer, ByRef Str_Salida As String)
'--------------------------------------------------------------------------------
'  Descripci¢n: devuelve debeCod o haberCod dependiendo si es debe o haber.
'  Autor: FGZ
'  Fecha: 26/10/2004
'-------------------------------------------------------------------------------
Dim cadena As String

    If Debe Then
        cadena = debeCod
    Else
        cadena = haberCod
    End If
    
    If Completar Then
        If Len(cadena) < Longitud Then
            cadena = String(Longitud - Len(cadena), "0") & cadena
        End If
    End If
    Str_Salida = cadena

End Sub

Private Sub ImporteTotal(ByVal Completar As Boolean, ByVal Longitud As Integer, ByRef Str_Salida As String)
'--------------------------------------------------------------------------------
'  Descripci¢n: devuelve el importe total, con dos decimales seguidos y
'               el separador de decimales es el definido en el modelo.
'  Autor: FGZ
'  Fecha: 26/10/2004
'-------------------------------------------------------------------------------
Dim I As Integer
Dim cadena As String
Dim Parte_Entera As Long
Dim Parte_Decimal As Integer
Dim Numero

    Numero = Split(CStr(totalImporte), ".")
    Parte_Entera = Fix(totalImporte)
    Parte_Decimal = IIf(Round((totalImporte - Parte_Entera) * 100, 0) > 0, Round((totalImporte - Parte_Entera) * 100, 0), 0)

    Numero(0) = Parte_Entera
    'cadena = Format(Parte_Entera, String(Longitud - 3, "0")) & SeparadorDecimales & Format(Parte_Decimal, "00")
    cadena = Format(Parte_Entera, String(Longitud - 3, "0")) & SeparadorDecimales
    If UBound(Numero) > 0 Then
        Numero(1) = Parte_Decimal
        cadena = cadena & Left(Numero(1) & "00", 2)
    Else
        cadena = cadena & "00"
    End If
    
    cadena = Replace(cadena, ",", ".")
    
    If Completar Then
        If Len(cadena) < Longitud Then
            cadena = String(Longitud - Len(cadena), "0") & cadena
        End If
    End If
    Str_Salida = cadena

End Sub

Private Sub totalRegistros(ByVal Total As Long, ByVal Completar As Boolean, ByVal Longitud As Integer, ByRef Str_Salida As String)
'--------------------------------------------------------------------------------
'  Descripci¢n: .
'  Autor: FGZ
'  Fecha: 26/10/2004
'-------------------------------------------------------------------------------
Dim I As Integer
Dim cadena As String

    cadena = Format(Total, String(Longitud, "0"))
        
    Str_Salida = cadena
End Sub



Private Sub Importe_Format(ByVal Monto As Single, ByVal Debe As Boolean, ByVal Completar As Boolean, ByVal Longitud As Integer, ByRef Str_Salida As String, ByVal Signo As String, ByVal Separador As String)
'--------------------------------------------------------------------------------
'Descripci¢n: devuelve el importe de la linea en el siguiente formato:
'               Si va al debe es + y - sino, con dos decimales seguidos con el
'               Separador
'Autor: FGZ
'Fecha: 25/04/2005
'-------------------------------------------------------------------------------
Dim I As Integer
Dim cadena As String
Dim Aux_Cadena As String
Dim Parte_Entera
Dim Parte_Decimal
Dim Numero

    Numero = Split(CStr(Monto), ".")
    Parte_Entera = Fix(Monto)
    Parte_Decimal = IIf(Round((Monto - Parte_Entera) * 100, 0) > 0, Round((Monto - Parte_Entera) * 100, 0), 0)
    
    Numero(0) = Parte_Entera
    If Debe Then
        cadena = " " & Format(Numero(0), String(Longitud - 3, "0"))
    Else
        cadena = "-" & Format(Numero(0), String(Longitud - 3, "0"))
    End If
    If UBound(Numero) > 0 Then
        Numero(1) = Parte_Decimal
        cadena = cadena & Left(Numero(1) & "00", 2)
    Else
        cadena = cadena & "00"
    End If
    
    If Debe Then
        Aux_Cadena = " " & Format(Numero(0), String(Longitud - 3, "0")) & Separador
    Else
        Aux_Cadena = "-" & Format(Numero(0), String(Longitud - 3, "0")) & Separador
    End If
    If UBound(Numero) > 0 Then
        Aux_Cadena = Aux_Cadena & Left(Numero(1) & "00", 2)
    Else
        Aux_Cadena = Aux_Cadena & "00"
    End If
    Aux_Cadena = Replace(Aux_Cadena, ",", ".")
    totalImporte = totalImporte + Abs(CSng(Aux_Cadena))
    
    If Completar Then
        If Len(cadena) < Longitud Then
            cadena = String(Longitud - Len(cadena), "0") & cadena
        End If
    End If
    Str_Salida = cadena

End Sub



