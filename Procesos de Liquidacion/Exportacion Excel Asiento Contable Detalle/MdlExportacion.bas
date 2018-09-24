Attribute VB_Name = "MdlExportacion"
Option Explicit

Global Const Version = "1.06"
Global Const FechaVersion = "17/10/2014"        ' Sebastian Stremel
Global Const UltimaModificacion = " "           ' CAS-27413 - NGA - Iron Mountain - Agregar columna Exportación análisis detallado
Global Const UltimaModificacion1 = " "          ' Se agrego columna descripcion de la cuenta a la exportacion.


'Global Const Version = "1.05"
'Global Const FechaVersion = "27/05/2014"        ' LED
'Global Const UltimaModificacion = " "           ' CAS-24599 - Tabacal - Reporte detallado asiento
'Global Const UltimaModificacion1 = " "          ' Se agrego columna cantidad a la exportacion.

'Global Const Version = "1.04"
'Global Const FechaVersion = "27/05/2014"       ' Borrelli Facundo
'Global Const UltimaModificacion = " "          ' Se libera la version 1.04, que soluciona parte del CAS-25024 - TSFOT - BUG EN EXPORTAR ASIENTO CONTABLE
'Global Const UltimaModificacion1 = " "         ' Se modificó el nombre del archivo que genera Detalle_asi-mm-yyyy-proceso.csv

'Global Const Version = "1.03"
'Global Const FechaVersion = "26/09/2012"       ' FGZ
'Global Const UltimaModificacion = " "                ' Se valida que exista la carpeta \PorUsr y \usuario, si no existen las crea.
'Global Const UltimaModificacion1 = " "              ' CAS-13764 - H&A - Visualizacion de Archivos Externos
'                                               Se le agregó sub control de versiones


'Const Version = "1.02"
'Const FechaVersion = "18/08/2009"
'Global Const UltimaModificacion = " Manuel Lopez"
'Global Const UltimaModificacion1 = "Encriptacion de string connection "


'Const Version = "1.01"
'Const FechaVersion = "24/02/2009"
''Martin Ferraro - Mas logs


'---------------------------------------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------------------------------
Private Type TR_Datos_Varios
    Convenio_Lecop As String        'String   long 4  -
    Filler As String                'String   long 1  -
    Cliente_Ya_Existente As String  'String   long 1  -
End Type

Global IdUser As String
Global Fecha As Date
Global hora As String

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
    Flog.writeline "Version                  : " & Version
    Flog.writeline "Fecha Ultima Modificacion: " & FechaVersion
    Flog.writeline "Modificacion = " & UltimaModificacion
    Flog.writeline "Modificacion = " & UltimaModificacion1
    Flog.writeline "-------------------------------------------------"
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 0) & "PID = " & PID
    
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
    
    Version_Valida = ValidarV(Version, 66, TipoBD)
    If Not Version_Valida Then
        'Actualizo el progreso
        MyBeginTrans
            StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcprogreso = 0, bprcestado = 'Error de Version', bprcpid = " & PID & " WHERE bpronro = " & NroProcesoBatch
            objconnProgreso.Execute StrSql, , adExecuteNoRecords
        MyCommitTrans
        Flog.writeline
        GoTo Fin
    End If
    
    On Error GoTo ME_Main

    'Cambio el estado del proceso a Procesando y el PID
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcprogreso = 0, bprcestado = 'Procesando', bprcpid = " & PID & " WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    Flog.writeline Espacios(Tabulador * 0) & "Archivo a Generar = " & Nombre_Arch
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE btprcnro = 66 AND bpronro =" & NroProcesoBatch
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
    
Fin:
    Flog.Close
    If objConn.State = adStateOpen Then objConn.Close
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
    GoTo Fin:
End Sub

Public Sub Generacion(ByVal bpronro As Long, ByVal Vol_Cod As Long, ByVal Separador As String)
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
Dim directorio As String
Dim Archivo As String
Dim Intentos As Integer
Dim carpeta

Dim NroLiq As Long
Dim strLinea As String
Dim Aux_Linea As String
Dim Texto As String

'Registros
Dim rs_Modelo As New ADODB.Recordset
Dim rs_Detalles As New ADODB.Recordset
Dim rs_Periodo As New ADODB.Recordset
Dim rs_mod_asiento As New ADODB.Recordset

'Archivo de exportacion
StrSql = "SELECT sis_dirsalidas FROM sistema WHERE sisnro = 1 "
OpenRecordset StrSql, rs
If Not rs.EOF Then
    directorio = Trim(rs!sis_dirsalidas)
End If

'StrSql = "SELECT * FROM modelo WHERE modnro = 234"
'OpenRecordset StrSql, rs_Modelo
'If Not rs_Modelo.EOF Then
'    If Not IsNull(rs_Modelo!modarchdefault) Then
'        Directorio = Directorio & Trim(rs_Modelo!modarchdefault)
'        Flog.writeline Espacios(Tabulador * 0) & "Directorio de generacion = " & Directorio
'    Else
'        Flog.writeline Espacios(Tabulador * 1) & "El modelo no tiene configurada la carpeta desteino. El archivo será generado en el directorio default"
'    End If
'Else
'    Flog.writeline Espacios(Tabulador * 1) & "No se encontró el modelo. El archivo será generado en el directorio default"
'End If


StrSql = "SELECT * FROM modelo WHERE modnro = 234"
OpenRecordset StrSql, rs_Modelo
If Not rs_Modelo.EOF Then
    If Not IsNull(rs_Modelo!modarchdefault) Then
        'VALIDO QUE EXISTA LA RUTA
        'directorio = directorio & Trim(rs_Modelo!modarchdefault)
        directorio = ValidarRuta(directorio, "\PorUsr", 1)
        directorio = ValidarRuta(directorio, "\" & IdUser, 1)
        directorio = ValidarRuta(directorio, "\" & Trim(rs_Modelo!modarchdefault), 1)
    Else
        Flog.writeline Espacios(Tabulador * 1) & "El modelo no tiene configurada la carpeta desteino. El archivo será generado en el directorio default"
    End If
Else
    Flog.writeline Espacios(Tabulador * 1) & "No se encontró el modelo. El archivo será generado en el directorio default"
End If

    
'cargo el periodo
StrSql = "SELECT * FROM proc_vol "
StrSql = StrSql & " INNER JOIN periodo ON proc_vol.pliqnro = periodo.pliqnro "
StrSql = StrSql & " WHERE vol_cod = " & Vol_Cod
OpenRecordset StrSql, rs_Periodo
If rs_Periodo.EOF Then
    Flog.writeline Espacios(Tabulador * 1) & "No se encontró el Periodo"
    Exit Sub
Else
     NroLiq = rs_Periodo!pliqnro
End If

'Seteo el nombre del archivo generado
'FGZ - 01/03/2013 ------------------------------------------------
'Se le agregó el nro de batch_proceso al nombre del archivo
'Archivo = directorio & "\Detalle_asi_" & Format(CStr(Month(rs_Periodo!pliqmes)), "00") & Format(CStr(rs_Periodo!pliqanio), "0000") & ".csv"
Archivo = directorio & "\Detalle_asi_" & Format(CStr(Month(rs_Periodo!pliqmes)), "00") & Format(CStr(rs_Periodo!pliqanio), "0000") & "-" & NroProcesoBatch & ".csv"
'FGZ - 01/03/2013 ------------------------------------------------

Set fs = CreateObject("Scripting.FileSystemObject")
'Activo el manejador de errores
On Error Resume Next
Set fExport = fs.CreateTextFile(Archivo, True)
If Err.Number <> 0 Then
    Flog.writeline Espacios(Tabulador * 1) & "La carpeta Destino no existe. Se creará."
    Set carpeta = fs.CreateFolder(directorio)
    Set fExport = fs.CreateTextFile(Archivo, True)
    
Else
   Flog.writeline Espacios(Tabulador * 1) & "Se creo el archivo en destino."
 End If

'desactivo el manejador de errores
On Error GoTo 0


' Comienzo la transaccion
MyBeginTrans

'Busco los procesos a evaluar
StrSql = "SELECT detalle_asi.*, mod_asiento.masidesc  FROM  detalle_asi "
StrSql = StrSql & " INNER JOIN mod_asiento ON mod_asiento.masinro = detalle_asi.masinro "
StrSql = StrSql & " WHERE detalle_asi.vol_cod =" & Vol_Cod
StrSql = StrSql & " ORDER BY detalle_asi.masinro, detalle_asi.ternro, detalle_asi.dldescripcion "
OpenRecordset StrSql, rs_Detalles

'seteo de las variables de progreso
Progreso = 0
CConceptosAProc = rs_Detalles.RecordCount
If CConceptosAProc = 0 Then
    CConceptosAProc = 1
    Flog.writeline Espacios(Tabulador * 1) & " No hay Detalles para ese Proceso de Volcado " & Vol_Cod
End If
IncPorc = (100 / CConceptosAProc)

'Procesamiento
If rs_Detalles.EOF Then
    Flog.writeline Espacios(Tabulador * 2) & "No hay nada que procesar"
End If
Flog.writeline Espacios(Tabulador * 2) & "Registros a Exportar: " & CConceptosAProc
        'Genero los encabezados
        Aux_Linea = "Detalle de lineas de asiento"
        fExport.writeline Aux_Linea
        fExport.writeline ""
        
        Aux_Linea = "Modelo" & Separador & "Descripción" & Separador & "Cuenta" & Separador & "Descripcion Cuenta" & Separador & "Cantidad" & Separador & "Monto" & Separador & "Acumulado" & Separador & "Porcentaje" & Separador & "Legajo" & Separador & "Apellido" & Separador & "Proyecto" & Separador & "Nivel Costo 1" & Separador & "Nivel Costo 2" & Separador & "Nivel Costo 3" & Separador & "Nivel Costo 4" & Separador & "Tipo Origen" & Separador & "Origen"
        fExport.writeline Aux_Linea


Do While Not rs_Detalles.EOF

        'Modelo, Descripcion,cuenta,monto
        Aux_Linea = rs_Detalles!masinro & " - " & rs_Detalles!masidesc & Separador & rs_Detalles!dldescripcion & Separador & rs_Detalles!Cuenta & Separador & rs_Detalles!Linadesc
        
        'cantidad, monto
        Aux_Linea = Aux_Linea & Separador & Format(rs_Detalles!dlcantidad, "########0.00") & Separador & Format(rs_Detalles!dlmonto, "########0.00")
        
        'Monto acumulado, Porcentaje, Legajo
        Aux_Linea = Aux_Linea & Separador & Format(rs_Detalles!dlmontoacum, "########0.00") & Separador & Format(rs_Detalles!dlporcentaje, "########0.00") & Separador & rs_Detalles!empleg
        
        'Apellido, Proyecto
        Aux_Linea = Aux_Linea & Separador & rs_Detalles!terape & Separador & rs_Detalles!proyecto
        
        'NivelCosto1
        Aux_Linea = Aux_Linea & Separador & IIf(EsNulo(rs_Detalles!dlcosto1), "0", rs_Detalles!dlcosto1)
        
        'NivelCosto2
        Aux_Linea = Aux_Linea & Separador & IIf(EsNulo(rs_Detalles!dlcosto2), "0", rs_Detalles!dlcosto2)
        
        'NivelCosto3
        Aux_Linea = Aux_Linea & Separador & IIf(EsNulo(rs_Detalles!dlcosto3), "0", rs_Detalles!dlcosto3)
        
        'NivelCosto4
        Aux_Linea = Aux_Linea & Separador & IIf(EsNulo(rs_Detalles!dlcosto4), "0", rs_Detalles!dlcosto4)
        
        'Tipo Origen
        Aux_Linea = Aux_Linea & Separador & IIf(rs_Detalles!tipoOrigen = "1", " Concepto ", " Acumulador ")
        
        'Origen
        Aux_Linea = Aux_Linea & Separador & rs_Detalles!Origen
        
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
    rs_Detalles.MoveNext
Loop
'Cierro el archivo creado
fExport.Close

'Fin de la transaccion
MyCommitTrans

Flog.writeline Espacios(Tabulador * 2) & "Termino de Exportar "

Fin:
If rs_Detalles.State = adStateOpen Then rs_Detalles.Close
If rs_Modelo.State = adStateOpen Then rs_Modelo.Close
If rs_Periodo.State = adStateOpen Then rs_Periodo.Close
If rs_mod_asiento.State = adStateOpen Then rs_mod_asiento.Close

Set rs_Detalles = Nothing
Set rs_Modelo = Nothing
Set rs_Periodo = Nothing
Set rs_mod_asiento = Nothing

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
Dim Vol_Cod As Long
Dim Periodo As Long
Dim aux As String

'Orden de los parametros
'Periodo
'proceso de volcado
'Separador de campos en el archivo a generar

Separador = "@"
' Levanto cada parametro por separado
If Not IsNull(parametros) Then
    If Len(parametros) >= 1 Then
    
        pos1 = 1
        pos2 = InStr(pos1, parametros, Separador) - 1
        Vol_Cod = Mid(parametros, pos1, pos2 - pos1 + 1)
        
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
Call Generacion(bpronro, Vol_Cod, Separador_de_Campos)
End Sub



Public Function ValidarV(ByVal Version As String, ByVal TipoProceso As Long, ByVal TipoBD As Integer) As Boolean
' ---------------------------------------------------------------------------------------------
' Descripcion: Funcion que determina si el proceso esta en condiciones de ejecutarse.
' Autor      : FGZ
' Fecha      : 05/08/2009
' ---------------------------------------------------------------------------------------------
Dim V As Boolean
Dim Texto As String
Dim rs As New ADODB.Recordset

On Error GoTo ME_Version

V = True

Select Case TipoProceso
Case 66: 'Exportacion de detalle de Asiento Contable
'    If Version >= "1.03" Then
'
'
'        Texto = "Revisar que exista tabla XXXX y su estructura sea correcta."
'
'        StrSql = "Select campo FROM tabla WHERE campo = 1"
'        OpenRecordset StrSql, rs
'
'        V = True
'    End If
    



'Cambio en busqueda de embargos bus_embargos
Case Else:
    Texto = "version correcta"
    V = True
End Select



    ValidarV = V
Exit Function

ME_Version:
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 1) & "Estructura de BD incompatible con la version del proceso."
    Flog.writeline Espacios(Tabulador * 1) & Texto
    Flog.writeline
    V = False
End Function


