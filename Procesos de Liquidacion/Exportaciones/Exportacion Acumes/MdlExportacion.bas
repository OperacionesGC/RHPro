Attribute VB_Name = "MdlExportacion"
Option Explicit

'Global Const Version = "1.01"
'Global Const FechaModificacion = "01/08/2008"
'Global Const UltimaModificacion = " "
'   01/08/2008 - Fernando Favre - CUSTOM - Arlei - Se agrego la posibilidad de generar el archivo en .../In-OutPorUsr/../<user>
'                Se esta manera cada reporte generado no es compartido por el resto de usuarios.
'                El manejo de la seguridad de los directorios queda en manos del administrador de la empresa

'Global Const Version = 1.02
'Global Const FechaModificacion = "19/08/2009"   'Encriptacion de string connection
'Global Const UltimaModificacion = "Manuel Lopez"
'Global Const UltimaModificacion1 = "Encriptacion de string connection"


'Global Const Version = 1.03
'Global Const FechaModificacion = "13/07/2011"   'Se agrega la barra que faltaba antes de la carpeta \PorUsr
'Global Const UltimaModificacion = "MED"
'Global Const UltimaModificacion1 = "Se agrega la barra que faltaba antes de la carpeta \PorUsr"

Global Const Version = 1.04
Global Const FechaModificacion = "09/01/2012"
Global Const UltimaModificacion = "Gonzalez Nicolás"
Global Const UltimaModificacion1 = "Se valida que exista la carpeta \PorUsr y \usuario, si no existen las crea."


Global IdUser As String
Global Fecha As Date
Global hora As String

Global Fecha_Inicio_periodo As Date
Global Fecha_Fin_Periodo As Date
Global StrSql2 As String


Public Sub Main()
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento inicial del Proceso.
' Autor      : Fernando Favre
' Fecha      : 10/02/2005
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
    
    Nombre_Arch = PathFLog & "Exp_Acum_Mensuales" & "-" & NroProcesoBatch & ".log"
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Nombre_Arch, True)
    
    Flog.writeline "Inicio Proceso de Exportación Acumulados Mensuales: " & Now
    Flog.writeline "Cambio el estado del proceso a Procesando"
    
    ' Obtengo el Process ID
    PID = GetCurrentProcessId
    Flog.writeline "-----------------------------------------------------------------"
    Flog.writeline "Version = " & Version
    Flog.writeline "Modificacion = " & UltimaModificacion
    Flog.writeline "Modificacion = " & UltimaModificacion1
    Flog.writeline "Fecha = " & FechaModificacion
    Flog.writeline "-----------------------------------------------------------------"
    Flog.writeline
    Flog.writeline "PID = " & PID
    
    'Cambio el estado del proceso a Procesando y el PID
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcprogreso = 0, bprcestado = 'Procesando', bprcpid = " & PID & " WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE btprcnro = 72 AND bpronro =" & NroProcesoBatch
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
    Flog.writeline "Tiempo del proceso (milisegundos): " & (TiempoFinalProceso - TiempoInicialProceso)
    
    If Not HuboError Then
        StrSql = "UPDATE batch_proceso SET bprctiempo ='100', bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado' WHERE bpronro = " & NroProcesoBatch
    Else
        StrSql = "UPDATE batch_proceso SET bprctiempo ='100', bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Error' WHERE bpronro = " & NroProcesoBatch
    End If
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    Flog.Close
    objconnProgreso.Close
    objConn.Close
    
End Sub
Public Sub Generacion(ByVal ammes As Long, ByVal amanio As Long, ByVal acuNro As String, ByVal opternro As Byte, ByVal Ternro As Long, ByVal bpronro As Long)
' --------------------------------------------------------------------------------------------
' Descripcion: Procedimiento de generacion de la Exportacion de Acumuladores Mensuales
' Autor      : Fernando Favre
' Fecha      : 11/02/2005
' Ult. Mod   :
' Fecha      :
' --------------------------------------------------------------------------------------------
Dim Detalle As String

Dim rs_Acu_mes As New ADODB.Recordset
Dim rs_Modelo As New ADODB.Recordset
Dim rs_Empleado As New ADODB.Recordset

Const ForReading = 1
Const TristateFalse = 0
Dim fExportDet
Dim Directorio As String
Dim Archivo As String
Dim carpeta

Dim Cantidad As Integer
Dim cantidadProcesada As Integer

Dim sql As String
Dim listempleados As String
Dim GeneroArchivo

GeneroArchivo = True

'Archivo de exportacion
StrSql = "SELECT sis_dirsalidas FROM sistema WHERE sisnro = 1 "
OpenRecordset StrSql, rs
If Not rs.EOF Then
    Directorio = Trim(rs!sis_dirsalidas)
End If

StrSql = "SELECT modarchdefault FROM modelo WHERE modnro = 250"
OpenRecordset StrSql, rs_Modelo
If Not rs_Modelo.EOF Then
    If Not IsNull(rs_Modelo!modarchdefault) Then
        'VALIDO QUE EXISTA LA RUTA
        Directorio = ValidarRuta(Directorio, "\PorUsr", 1)
        Directorio = ValidarRuta(Directorio, "\" & IdUser, 1)
        Directorio = ValidarRuta(Directorio, "\" & Trim(rs_Modelo!modarchdefault), 1)
        'Directorio = Directorio & "\PorUsr\" & IdUser & Trim(rs_Modelo!modarchdefault)
    Else
        Flog.writeline "El modelo no tiene configurada la carpeta destino. El archivo será generado en el directorio default"
    End If
Else
    Flog.writeline "No se encontró el modelo. El archivo será generado en el directorio default"
End If
rs_Modelo.Close


'Activo el manejador de errores
On Error Resume Next

'Archivo para el detalle del Pedido de Pago
Archivo = Directorio & "\acu_mensuales.txt"
Set fs = CreateObject("Scripting.FileSystemObject")
Set fExportDet = fs.CreateTextFile(Archivo, True)

'If Err.Number <> 0 Then
'    Flog.writeline "La carpeta Destino no existe. Se creará."
'   Set carpeta = fs.CreateFolder(Directorio)
'    Set fExportDet = fs.CreateTextFile(Archivo, True)
'    GeneroArchivo = False
'End If
'desactivo el manejador de errores

On Error GoTo CE

' Comienzo la transaccion
MyBeginTrans

StrSql = "SELECT acu_mes.*, empleado.empleg "
StrSql = StrSql & " FROM acu_mes "
StrSql = StrSql & " INNER JOIN empleado ON acu_mes.ternro = empleado.ternro "
StrSql = StrSql & " WHERE acu_mes.ammes = " & ammes
StrSql = StrSql & " AND acu_mes.amanio = " & amanio
StrSql = StrSql & " AND acu_mes.acunro IN (" & acuNro & ")"

If opternro = 1 Then
    StrSql = StrSql & " AND acu_mes.ternro = " & Ternro
End If
If opternro = 2 Then
    sql = "SELECT ternro FROM batch_empleado WHERE bpronro = " & bpronro
    OpenRecordset sql, rs_Empleado
    listempleados = ""
    Do While Not rs_Empleado.EOF
        listempleados = listempleados & "," & rs_Empleado!Ternro
        rs_Empleado.MoveNext
    Loop
    listempleados = Mid(listempleados, 2, Len(listempleados) - 1)
    rs_Empleado.Close
    Set rs_Empleado = Nothing
    StrSql = StrSql & " AND acu_mes.ternro IN ( " & listempleados & ")"
End If

StrSql = StrSql & " ORDER BY acu_mes.amanio, acu_mes.ammes, acu_mes.ternro, acu_mes.acunro"
OpenRecordset StrSql, rs_Acu_mes

Cantidad = rs_Acu_mes.RecordCount
cantidadProcesada = Cantidad

'seteo de las variables de progreso
Progreso = 0
CConceptosAProc = rs_Acu_mes.RecordCount
If CConceptosAProc = 0 Then
    CConceptosAProc = 1
    Flog.writeline Espacios(Tabulador * 1) & " No hay acumulados"
        GeneroArchivo = False
End If
IncPorc = (99 / CConceptosAProc)
'Procesamiento
If rs_Acu_mes.EOF Then
    Flog.writeline Espacios(Tabulador * 2) & "No hay nada que procesar"
    GeneroArchivo = False
End If


Do While Not rs_Acu_mes.EOF
           
    Detalle = rs_Acu_mes!empleg & ";" & rs_Acu_mes!acuNro & ";" & rs_Acu_mes!ammonto & ";"
    Detalle = Detalle & rs_Acu_mes!amcant & ";" & rs_Acu_mes!amanio & ";" & rs_Acu_mes!ammes
    fExportDet.writeline Detalle
                
    TiempoAcumulado = GetTickCount
          
          
    cantidadProcesada = cantidadProcesada - 1
          
    'Actualizo el progreso del Proceso
    Progreso = Progreso + IncPorc
    TiempoAcumulado = GetTickCount
    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & _
             ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & _
             "' WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    rs_Acu_mes.MoveNext
Loop
 If HuboError = False And Err.Number = 0 And GeneroArchivo = True Then
 
 Flog.writeline "========================================================================================================="
 Flog.writeline Espacios(Tabulador * 2) & " "
 Flog.writeline " Se ha generado el Archivo acu_mensuales.txt en el directorio: "
 Flog.writeline Espacios(Tabulador * 2) & Archivo
 Flog.writeline Espacios(Tabulador * 2) & " "
 Flog.writeline "=========================================================================================================="
 End If
rs_Acu_mes.Close
fExportDet.Close

MyCommitTrans

Set rs_Acu_mes = Nothing
Set rs_Modelo = Nothing

Exit Sub

CE:
    Flog.writeline "================================================================="
    HuboError = True
    Flog.writeline " Error: " & Err.Description & Now
    Flog.writeline "================================================================="

End Sub

Public Sub LevantarParamteros(ByVal bpronro As Long, ByVal parametros As String)
' --------------------------------------------------------------------------------------------
' Descripcion: Procedimiento para levantar los parametros pasados en batch_proceso en bprcparam
' Autor      : Fernando Favre
' Fecha      : 10/02/2005
' Ult. Mod   :
' Fecha      :
' --------------------------------------------------------------------------------------------
Dim pos1 As Integer
Dim pos2 As Integer
Dim Separador As String

Dim ArrParametros
Dim ammes As Long
Dim amanio As Long
Dim acuNro As String
Dim opternro As Byte
Dim Ternro As Long

'Orden de los parametros
'Mes
'Año
'Acumulador en la lista que se haya elegido de acumuladores
'Opcion Empleado (1 Uno en particular, 2 Filtro, 3 Todos)
'ternro del Empleado en caso que se haya elegido uno

ArrParametros = Split(parametros, "@")
' Levanto cada parametro por separado
ammes = ArrParametros(0)
amanio = ArrParametros(1)
acuNro = ArrParametros(2)
opternro = ArrParametros(3)
If opternro = 1 Then
    Ternro = ArrParametros(4)
End If

Call Generacion(ammes, amanio, acuNro, opternro, Ternro, bpronro)

End Sub




