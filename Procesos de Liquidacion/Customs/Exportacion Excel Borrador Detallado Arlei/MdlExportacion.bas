Attribute VB_Name = "MdlExportacion"
Option Explicit

Global Const Version = "1.01"
Global Const FechaModificacion = "01/08/2008"
Global Const UltimaModificacion = " "
'   01/08/2008 - Fernando Favre - CUSTOM - Arlei - Se agrego la posibilidad de generar el archivo en .../In-OutPorUsr/<user>/..
'                Se esta manera cada reporte generado no es compartido por el resto de usuarios.
'                El manejo de la seguridad de los directorios queda en manos del administrador de la empresa


Private Type TR_Datos_Varios
    Convenio_Lecop As String        'String   long 4  -
    Filler As String                'String   long 1  -
    Cliente_Ya_Existente As String  'String   long 1  -
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
    
    Nombre_Arch = PathFLog & "Exp_Borrador_Detallado" & "-" & NroProcesoBatch & ".log"
    
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
    
    'Cambio el estado del proceso a Procesando y el PID
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcprogreso = 0, bprcestado = 'Procesando', bprcpid = " & PID & " WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE btprcnro = 67 AND bpronro =" & NroProcesoBatch
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

Public Sub Generacion(ByVal bpronro As Long, ByVal Proceso As Long, ByVal Separador As String)
' --------------------------------------------------------------------------------------------
' Descripcion: Procedimiento de generacion del archivo excel del borrador detallado.
' Autor      : FGZ
' Fecha      : 16/12/2004
' Ult. Mod   :
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

Dim Te1 As Boolean
Dim Te2 As Boolean
Dim Te3 As Boolean
Dim Encabezado As String
Dim Aux_Nombre As String

'Registros
Dim rs_BorradorDeta As New ADODB.Recordset
Dim rs_BorradorDeta_Det As New ADODB.Recordset
Dim rs_Cantidad As New ADODB.Recordset
Dim rs_Modelo As New ADODB.Recordset

'Archivo de exportacion
StrSql = "SELECT sis_dirsalidas FROM sistema WHERE sisnro = 1 "
OpenRecordset StrSql, rs
If Not rs.EOF Then
    Directorio = Trim(rs!sis_dirsalidas)
End If

StrSql = "SELECT * FROM modelo WHERE modnro = 244"
OpenRecordset StrSql, rs_Modelo
If Not rs_Modelo.EOF Then
    If Not IsNull(rs_Modelo!modarchdefault) Then
        Directorio = Directorio & "PorUsr\" & IdUser & Trim(rs_Modelo!modarchdefault)
'        Directorio = Directorio & Trim(rs_Modelo!modarchdefault)
    Else
        Flog.writeline Espacios(Tabulador * 1) & "El modelo no tiene configurada la carpeta desteino. El archivo será generado en el directorio default"
    End If
Else
    Flog.writeline Espacios(Tabulador * 1) & "No se encontró el modelo. El archivo será generado en el directorio default"
End If

'cargo el periodo
StrSql = "SELECT * FROM rep_borradordeta "
StrSql = StrSql & " WHERE bpronro = " & Proceso
StrSql = StrSql & " ORDER BY orden"
OpenRecordset StrSql, rs_BorradorDeta
If rs_BorradorDeta.EOF Then
    Flog.writeline Espacios(Tabulador * 1) & "No se encontró el Periodo"
    Exit Sub
End If

'Seteo el nombre del archivo generado
Archivo = Directorio & "\Borrador_Det_" & rs_BorradorDeta!pliqdesc & "_Proceso_" & Proceso & ".csv"
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
MyBeginTrans

'Para calcular el progreso
StrSql = "SELECT * FROM rep_borrdeta_det "
StrSql = StrSql & " WHERE rep_borrdeta_det.bpronro = " & Proceso
OpenRecordset StrSql, rs_Cantidad

'seteo de las variables de progreso
Progreso = 0
CConceptosAProc = rs_Cantidad.RecordCount
If CConceptosAProc = 0 Then
    CConceptosAProc = 1
End If
IncPorc = (100 / CConceptosAProc)

'Procesamiento
If rs_Cantidad.EOF Then
    Flog.writeline Espacios(Tabulador * 2) & "No hay nada que procesar"
End If
If rs_Cantidad.State = adStateOpen Then rs_Cantidad.Close

'Genero los encabezados
Aux_Linea = "CONTROL LIQUIDACIÓN SUELDOS"
fExport.writeline Aux_Linea
fExport.writeline ""

Te1 = False
Te2 = False
Te3 = False

Aux_Linea = "Empleado" & Separador & "Apellido y Nombre" & Separador & "Período" & Separador & "Proceso" & Separador & "Depto." & Separador & "Categoría" & Separador & "Ingreso" & Separador & "Cuil" & Separador & "Contrato" & Separador & "Código" & Separador & "Concepto" & Separador & "Cantidad" & Separador & "Monto"
If Not rs_BorradorDeta.EOF Then
    If Not EsNulo(rs_BorradorDeta!tedabr3) Then
        Aux_Linea = rs_BorradorDeta!tedabr3 & Separador & Aux_Linea
        Te3 = True
    End If
    If Not EsNulo(rs_BorradorDeta!tedabr2) Then
        Aux_Linea = rs_BorradorDeta!tedabr2 & Separador & Aux_Linea
        Te2 = True
    End If
    If Not EsNulo(rs_BorradorDeta!tedabr1) Then
        Aux_Linea = rs_BorradorDeta!tedabr1 & Separador & Aux_Linea
        Te1 = True
    End If
End If
fExport.writeline Aux_Linea
        
Do While Not rs_BorradorDeta.EOF
    
    StrSql = "SELECT * FROM rep_borrdeta_det"
    StrSql = StrSql & " WHERE rep_borrdeta_det.bpronro = " & Proceso
    StrSql = StrSql & " AND rep_borrdeta_det.ternro = " & rs_BorradorDeta!ternro
    StrSql = StrSql & " AND rep_borrdeta_det.pronro = " & rs_BorradorDeta!pronro
    StrSql = StrSql & " ORDER BY conccod "
    OpenRecordset StrSql, rs_BorradorDeta_Det
    
    Do While Not rs_BorradorDeta_Det.EOF
        
        Aux_Linea = ""
        'Agregar las primeras columnas
        'Estructuras
        If Te1 Then
            Aux_Linea = Aux_Linea & rs_BorradorDeta!estrdabr1
        End If
        If Te2 Then
            Aux_Linea = Aux_Linea & IIf(Te1, Separador, "") & rs_BorradorDeta!estrdabr2
        End If
        If Te3 Then
            Aux_Linea = Aux_Linea & IIf(Te1 Or Te2, Separador, "") & rs_BorradorDeta!estrdabr3
        End If
        
        'Demas campos del empleado
        'Empleado    Apellido y Nombre   Período Proceso Depto.  Categoría   Ingreso Cuil    Contrato
        If Te1 Or Te2 Or Te3 Then
            Aux_Linea = Aux_Linea & Separador
        End If
        
        Aux_Nombre = " "
        Aux_Nombre = Aux_Nombre & IIf(Not EsNulo(rs_BorradorDeta!apellido), rs_BorradorDeta!apellido, "")
        Aux_Nombre = Aux_Nombre & IIf(Not EsNulo(rs_BorradorDeta!apellido2), " " & rs_BorradorDeta!apellido2, "")
        Aux_Nombre = Aux_Nombre & IIf(Not EsNulo(rs_BorradorDeta!nombre), " " & rs_BorradorDeta!nombre, "")
        Aux_Nombre = Aux_Nombre & IIf(Not EsNulo(rs_BorradorDeta!nombre2), " " & rs_BorradorDeta!nombre2, "")
        'Legaj, Apellido y Nombre
        Aux_Linea = Aux_Linea & rs_BorradorDeta!Legajo & Separador & Aux_Nombre
        
        'Periodo, Proceso
        Aux_Linea = Aux_Linea & Separador & rs_BorradorDeta!pliqdesc & Separador & rs_BorradorDeta!prodesc
        
        'Centro de costo, Categoria
        Aux_Linea = Aux_Linea & Separador & rs_BorradorDeta!centrocosto & Separador & rs_BorradorDeta!categoria
        
        'Ingreso, Cuil
        Aux_Linea = Aux_Linea & Separador & rs_BorradorDeta!fecalta & Separador & rs_BorradorDeta!Cuil
        
        'Contrato
        Aux_Linea = Aux_Linea & Separador & rs_BorradorDeta!contrato
        
        Encabezado = Aux_Linea
        
        'Agrego el detalle
        'Código  Concepto    Cantidad    Monto
        Aux_Linea = Aux_Linea & Separador & rs_BorradorDeta_Det!Conccod & Separador & rs_BorradorDeta_Det!concabr & Separador & rs_BorradorDeta_Det!dlicant & Separador & rs_BorradorDeta_Det!dlimonto
        
        ' ------------------------------------------------------------------------
        'Escribo en el archivo de texto
        fExport.writeline Aux_Linea
            
        'Actualizo el progreso del Proceso
        Progreso = Progreso + IncPorc
        TiempoAcumulado = GetTickCount
        StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & _
                 ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & _
                 "' WHERE bpronro = " & NroProcesoBatch
        objconnProgreso.Execute StrSql, , adExecuteNoRecords
                
        'Siguiente proceso
        rs_BorradorDeta_Det.MoveNext
    Loop
                    
    'Agregar las lineas de los acumuladores
    'Acumulador 1
    Aux_Linea = ""
    If Not EsNulo(rs_BorradorDeta!acumval1) Then
        Aux_Linea = Encabezado
        Aux_Linea = Aux_Linea & Separador & " " & Separador & rs_BorradorDeta!acumdesc1 & Separador & " " & Separador & rs_BorradorDeta!acumval1
        fExport.writeline Aux_Linea
    End If
    
    'Acumulador 2
    Aux_Linea = ""
    If Not EsNulo(rs_BorradorDeta!acumval2) Then
        Aux_Linea = Encabezado
        Aux_Linea = Aux_Linea & Separador & " " & Separador & rs_BorradorDeta!acumdesc2 & Separador & " " & Separador & rs_BorradorDeta!acumval2
        fExport.writeline Aux_Linea
    End If
                        
    'Acumulador 3
    Aux_Linea = ""
    If Not EsNulo(rs_BorradorDeta!acumval3) Then
        Aux_Linea = Encabezado
        Aux_Linea = Aux_Linea & Separador & " " & Separador & rs_BorradorDeta!acumdesc3 & Separador & " " & Separador & rs_BorradorDeta!acumval3
        fExport.writeline Aux_Linea
    End If
                        
    'Acumulador 4
    Aux_Linea = ""
    If Not EsNulo(rs_BorradorDeta!acumval4) Then
        Aux_Linea = Encabezado
        Aux_Linea = Aux_Linea & Separador & " " & Separador & rs_BorradorDeta!acumdesc4 & Separador & " " & Separador & rs_BorradorDeta!acumval4
        fExport.writeline Aux_Linea
    End If
                
    'Siguiente
    rs_BorradorDeta.MoveNext
Loop
'Cierro el archivo creado
fExport.Close

'Fin de la transaccion
MyCommitTrans


Fin:
If rs_BorradorDeta.State = adStateOpen Then rs_BorradorDeta.Close
If rs_Modelo.State = adStateOpen Then rs_Modelo.Close
If rs_BorradorDeta_Det.State = adStateOpen Then rs_BorradorDeta_Det.Close

Set rs_BorradorDeta = Nothing
Set rs_Modelo = Nothing
Set rs_BorradorDeta_Det = Nothing

Exit Sub
CE:
    HuboError = True
    MyRollbackTrans
    GoTo Fin
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

Dim Separador_de_Campos As String
Dim Proceso As Long
Dim aux

'Orden de los parametros
'bpronro
'Separador

Separador = "@"
' Levanto cada parametro por separado
If Not IsNull(parametros) Then
    If Len(parametros) >= 1 Then
    
        pos1 = 1
        pos2 = InStr(pos1, parametros, Separador) - 1
        Proceso = Mid(parametros, pos1, pos2 - pos1 + 1)
        
        pos1 = pos2 + 2
        pos2 = Len(parametros)
        aux = Mid(parametros, pos1, pos2 - pos1 + 1)
        If aux = 1 Then
            Separador_de_Campos = ","
        Else
            Separador_de_Campos = ";"
        End If
        
    End If
End If
Call Generacion(bpronro, Proceso, Separador_de_Campos)
End Sub

