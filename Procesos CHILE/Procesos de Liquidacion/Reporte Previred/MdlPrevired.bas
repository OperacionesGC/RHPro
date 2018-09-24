Attribute VB_Name = "MdlPrevired"
Option Explicit

'Const Version = 1.01
'Const FechaVersion = "21/01/2007"
'Autor = Diego Rosso

'Const Version = 1.02
'Const FechaVersion = "9/03/2007"
'Modificacion = Diego Rosso - Se agrego para que inserte los empleados que no tienen errores porque si
                             'encontraba un empleado con error abortaba todo el proceso
'Const Version = 1.03
'Const FechaVersion = "13/03/2007"
'Modificacion = Diego Rosso - Se agrego para que se pueda poner un valor constante en una columna por confrep.
                             'y se agregaron mas validaciones.


'Const Version = 1.04
'Const FechaVersion = "01/02/2008"
'Modificacion = Diego Rosso - Se cambio la Configuracion de la columna 17, se paso de Tipo de Estructura a AC/CO
'                             Se pasa todos los montos a Positivo usando ABS

'Const Version = 1.05
'Const FechaVersion = "12/09/2008"
'Modificacion = Javier Irastorza - Se corrigio la Configuracion de la columna 52, EN parte la consideraba como TE y en otras como AC/CO
'Modificacion = Javier Irastorza - Se corrigio la Configuracion de la columna 53, Cuando es Fonasa no se informa Cotizacion Obligatoria

'Const Version = 1.06
'Const FechaVersion = "19/09/2008"
'Modificacion = Javier Irastorza - Se corrigio exportacion de las fechas desde y hasta de los movimientos

'Const Version = 1.07
'Const FechaVersion = "26/09/2008"
''Modificacion = Javier Irastorza - Se agrego una linea para los distintos tipos de movimientos a informar


'Const Version = "1.08"
'Const FechaVersion = "06/10/2008"
''Modificacion = FGZ - Agregados de log


'Const Version = "1.09"
'Const FechaVersion = "20/11/2008"
'Modificacion = FGZ - cambios en la forma en que busca los movimientos
'                       Antes los sacaba de historico de estructuras
'                       ahora se sacan de fases y de licencias
'                           OJO!! igualmente las estructuras de tipos de movimientos deben tener configurados
'                                   los tipos de codigos y los tipos de licencias

'Const Version = "1.10"
'Const FechaVersion = "05/05/2009"
'Modificacion = Martin Ferraro - Se agrego parametro tnomina
'                                Se quitaron las condiciones que no permitian imprimir un registro cuando habia un error
'                                Campo tipo Pago es igual al tipo de nomina del filtro

'Const Version = "1.11"
'Const FechaVersion = "26/06/2009"
'Modificacion = Stankunas Cesar - Se modificó el modelo a la nueva versión (12.0 - 13 de Abrildel 2009) de 105 campos


'Const Version = "1.12"
'Const FechaVersion = "15/07/2009"
'Modificacion = Martin Ferraro - Abrir lineas por Movimientos, APVI y APVC para version 105 campos

'Const Version = "1.13"
'Const FechaVersion = "04/08/2009"
'Modificacion = MB - Correccion de errores de 105 campos, conexion encriptada

'Const Version = "1.14"
'Const FechaVersion = "07/08/2009"
'Modificacion = MB - Cambio de AFP de Estructura a concepto, arregloEstruc(26) por arreglo(26)

'Const Version = "1.15"
'Const FechaVersion = "12/08/2009"
'Modificacion = Martin Ferraro - Se quito formatnumber cuando suma en arreglo y agrego log

'Const Version = "1.16"
'Const FechaVersion = "13/08/2009"
'Modificacion = MB - Se cambio el tipo de linea en las lineas adicionales

'Const Version = "1.17"
'Const FechaVersion = "10/11/2009"
'Modificacion = MB - Se cambio el centro de costo arregloEstruc (105) para que tome 15 caracteres

'Const Version = "1.18"
'Const FechaVersion = "29/11/2010"
'Modificacion = Matias Dallegro Error formato se corrigen lineas de Gratificacion

'Const Version = "1.19"
'Const FechaVersion = "01/12/2010"
'Modificacion = MB se corrigen linea de Gratificacion

'Const Version = "1.21"
'Const FechaVersion = "28/12/2010"
''Martin Ferraro - En gratificaciones campo 84 renta impo ccaf estaba en cero, ahora arreglo(84)

'Const Version = "1.21"
'Const FechaVersion = "11/11/2011"
''Lisandro Moro - Correccion sql Gratificaciones.

'Const Version = "1.22"
'Const FechaVersion = "15/11/2011"
'Lisandro Moro - Correccion al cerrar los recorsets.

'Const Version = "1.23"
'Const FechaVersion = "20/12/2011"
'Sebastian Stremel - se realizan cambios en el tramo de asignacion familiar para deloitte chile

'Const Version = "1.24"
'Const FechaVersion = "02/02/2012"
'Sebastian Stremel - se amplia arreglo a 210 en vez de 110 y se realizan cambios para la lina 15, 100,101,102

'Const Version = "1.25"
'Const FechaVersion = "27/04/2012"
''Sebastian Stremel - si el empleado tiene 2 nombres y la suma de caracteres de esos 2 supera los 29 caracteres,
''entonces solo muestro el primero, debido a que previred no permite mostrar nombres cortados y el espacio tiene que ser de 30 lugares.


'Const Version = "1.26"
'Const FechaVersion = "31/05/2011"
'           FGZ - CAS-19856 - RHPro Consulting - ACS - Error Previred
'           Se agregó control por Licencia de codigo 11

'Const Version = "1.27"
'Const FechaVersion = "12/09/2013"
'           Sebastian Stremel - CAS-20959 - RHPro Consulting Chile - Cambio legal Previred
'           Si el proceso es finiquito agrego un movimiento con codigo 12

'Const Version = "1.28"
'Const FechaVersion = "28/10/2014"
'           Carmen Quintero - CAS-26972 - H&A - Bugs detectados en R4 - Error en el Mapeo de Documentos en el reporte Previred
'           Se agregó relacion con la tabla tipodocu_pais al momento de buscar el nro de documento de un empleado



'Const Version = "1.29"
'Const FechaVersion = "16/12/2014"
'                               Fernandez, Matias -CAS-28339 - RH PRO CHILE - Bug en Previred-
'                               Se genera una fila por cada concepto encontrado en la columna 44

'Const Version = "1.30"
'Const FechaVersion = "05/02/2014"
'                               Fernandez, Matias -CAS-28339 - RH PRO CHILE - Bug en Previred-
'                               Se genera una fila por cada concepto encontrado en la columna 44,
'                               solo hay una columna 44 del reporte si el concepto de la
' columna 40 que estoy mirando tambien esta en la columna 44 del confrep

'Const Version = "1.31"
'Const FechaVersion = "09/06/2015"
'                               Stremel,Sebastian - CAS-30835 - RH Pro Chile - Reporte previred-
'                               Se muestran solo los importes de los conceptos CCAF correspondiente a la empresa

'Const Version = "1.32"
'Const FechaVersion = "04/08/2015"
'                               Stremel,Sebastian - CAS-30835 - RH Pro Chile - Reporte previred-
'                               Se muestran solo los importes de los conceptos CCAF correspondiente a la empresa


'Const Version = "1.33"
'Const FechaVersion = "06/01/2016"
'                               06-01-2016 - MDF - CAS-34560 - RH Pro Chile- Bug en previred -
'                               Correccion en fechas de licencias a buscar.
 
 
 
Const Version = "1.34"
Const FechaVersion = "05/02/2016"
'                               06-02-2016 - MDF - CAS-34560 - RH Pro Chile- Bug en previred -
'                               se vuelve a corregir el limite de fechas en la cual se buscan las licencias
 

Private Type TregAPV
    Cod As Long
    Contrato As String
    FPago As Long
    Cotiza As Double
    Depositos As Double
End Type

'---------------------------------------------------------------
Global CantEmplError '08-03-07 Diego Rosso
Global CantEmplSinError '08-03-07 Diego Rosso
Global Errores As Boolean '08-03-07 Diego Rosso


Public Sub Main()
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento inicial del Generador de Reporte Previred.
' Autor      : Diego Rosso
' Fecha      : 21/01/2007
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------

Dim objconnMain As New ADODB.Connection
Dim strCmdLine
Dim Nombre_Arch As String
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

    On Error Resume Next
    'Abro la conexion
'    OpenConnection strconexion, objConn
'    If Err.Number <> 0 Then
'        Flog.writeline "Problemas en la conexion"
'        Exit Sub
'    End If
'    OpenConnection strconexion, objconnProgreso
'    If Err.Number <> 0 Then
'        Flog.writeline "Problemas en la conexion"
'        Exit Sub
'    End If
'    On Error GoTo 0
    
    On Error GoTo ME_Main
    
    Nombre_Arch = PathFLog & "Generacion_Previred" & "-" & NroProcesoBatch & ".log"
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Nombre_Arch, True)
    
    
    ' Obtengo el Process ID
    PID = GetCurrentProcessId
    Flog.writeline "-------------------------------------------------"
    Flog.writeline "Version                  : " & Version
    Flog.writeline "Fecha Ultima Modificacion: " & FechaVersion
    Flog.writeline "PID                      : " & PID
    Flog.writeline "-------------------------------------------------"
    Flog.writeline
    
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
    
    On Error GoTo ME_Main
    
    'Cambio el estado del proceso a Procesando
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcprogreso = 0 ,bprcestado = 'Procesando', bprcpid = " & PID & " WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    Flog.writeline "Pone el estado en procesando"
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE btprcnro = 156 AND bpronro =" & NroProcesoBatch
    OpenRecordset StrSql, rs_batch_proceso
    
    TiempoInicialProceso = GetTickCount
    
    If Not rs_batch_proceso.EOF Then
        bprcparam = rs_batch_proceso!bprcparam
        rs_batch_proceso.Close
        Set rs_batch_proceso = Nothing
        Call Previred(NroProcesoBatch, bprcparam)
    Else
        Flog.writeline "no encontró el proceso"
    End If
    
    TiempoFinalProceso = GetTickCount
    Flog.writeline
    Flog.writeline "**********************************************************"
    Flog.writeline
    Flog.writeline "Cantidad de Empleados Insertados: " & CantEmplSinError
    Flog.writeline "Cantidad de Empleados Con ERRORES: " & CantEmplError
    Flog.writeline
    Flog.writeline "Tiempo del proceso (milisegundos): " & (TiempoFinalProceso - TiempoInicialProceso)
    Flog.writeline
    Flog.writeline "**********************************************************"
    If Not Errores Then
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado', bprcprogreso = 100 WHERE bpronro = " & NroProcesoBatch
    Else
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Error', bprcprogreso = 100  WHERE bpronro = " & NroProcesoBatch
    End If
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
Fin:
    Flog.Close
    'If objConn.State = adStateOpen Then objConn.Close
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
    'MyBeginTrans
        StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcprogreso = 0, bprcestado = 'Error General', bprcpid = " & PID & " WHERE bpronro = " & NroProcesoBatch
        objconnProgreso.Execute StrSql, , adExecuteNoRecords
    'MyCommitTrans
End Sub


Public Sub Previred(ByVal bpronro As Long, ByVal Parametros As String)
' --------------------------------------------------------------------------------------------
' Descripcion: Procedimiento de generacion del reporte del Previred
' Autor      : Diego Rosso
' Fecha      : 20/01/2007
' Modifiacdo : 26/06/2009 - Stankunas Cesar - Se adecuo para el formato de 105 campos
' --------------------------------------------------------------------------------------------
Dim Empresa As Long
Dim Lista_Pro As String
Dim Fechadesde As Date
Dim Fechahasta As Date

Dim FechaDesdePeri As Date
Dim FechaHastaPeri As Date

Dim topeArreglo As Integer   'USAR ESTA VARIABLE PARA EL TOPE
Dim arreglo(210) As Double 'cambio valor a 210 antes 110
Dim arregloEstruc(210) As String 'antes 110
Dim tNomina As Integer

Dim arregloMov(30) As Integer
Dim arregloFecD(30) As Date
Dim arregloFecH(30) As Date
Dim total_mov As Integer
Dim aux As Integer

Dim arregloAPVI(30) As TregAPV
Dim arregloAPVC(30) As TregAPV
Dim total_APVI As Long
Dim total_APVC As Long

Dim I As Integer
Dim pos1 As Integer
Dim pos2 As Integer
Dim UltimoEmpleado As Long
Dim Apellido
Dim Apellido2
Dim NombreEmp
Dim RUT
Dim DV
Dim Num_linea
Dim Titulo
Dim Contador
Dim Sexo
Dim TipoPago
Dim EMPternro
Dim FUN
Dim EsFonasa
Dim SeguroCesantia As Boolean
Dim EstaIPS As Boolean
Dim cambioContrato As Boolean
Dim procesos

Dim Nacionalidad As Integer
Dim TipoLinea As String
Dim SolicSubsidioJoven As String

Dim sql As String
Dim estructuraContrato As Integer
'recordsets
Dim rs_Empleados As New ADODB.Recordset
Dim rs_Empleados1 As New ADODB.Recordset
Dim rs_CantEmpleados As New ADODB.Recordset
Dim rs_Acu_liq As New ADODB.Recordset
Dim rs_Empresa As New ADODB.Recordset
Dim rs_Confrep As New ADODB.Recordset
Dim rs_Conceptos As New ADODB.Recordset
Dim rs_Detliq As New ADODB.Recordset
Dim rs_Tercero As New ADODB.Recordset
Dim rs_Estructura As New ADODB.Recordset
Dim rs_Fases As New ADODB.Recordset
Dim rs_Aux As New ADODB.Recordset
Dim rs_Rut As New ADODB.Recordset
Dim rs_Estr_cod As New ADODB.Recordset
Dim rs_Nacionalidad As New ADODB.Recordset
Dim rs_impuniperi As New ADODB.Recordset
Dim rs_Contratos As New ADODB.Recordset
Dim dia As String
Dim mes As String
Dim Anio As String

Dim LongNombreEmp1 As Integer
Dim LongNombreEmp2 As Integer

Dim codCaja As Long
Dim teCajaCCAF As Integer
codCaja = 0
teCajaCCAF = 0
 ' Inicio codigo ejecutable
    On Error GoTo CE

' El formato de los parametros pasados es
'  (titulo del reporte, Todos_los_procesos, empresa,fechadesde, fechahasta)

' Levanto cada parametro por separado, el separador de parametros es "@"
Flog.writeline "Levantando Parametros  "
Flog.writeline Espacios(Tabulador * 1) & Parametros
Flog.writeline
If Not IsNull(Parametros) Then
    
    If Len(Parametros) >= 1 Then
        
        'TITULO
        '-----------------------------------------------------------
        pos1 = 1
        pos2 = InStr(pos1, Parametros, "@") - 1
        Titulo = Mid(Parametros, pos1, pos2)
        Flog.writeline "Posicion 1 = " & pos1
        Flog.writeline "Pos 2 = " & pos2
        Flog.writeline "Parametro Titulo = " & Titulo
        Flog.writeline
        '-------------------------------------------------------------
        
        'PROCESOS
        '-------------------------------------------------------------
        pos1 = pos2 + 2
        pos2 = InStr(pos1, Parametros, "@") - 1
        Lista_Pro = Mid(Parametros, pos1, pos2 - pos1 + 1)
        
        Flog.writeline "Posicion 1 = " & pos1
        Flog.writeline "Pos 2 = " & pos2
        Flog.writeline "Parametro Lista_Pro = " & Lista_Pro
        Flog.writeline
        ' esta lista tiene los nro de procesos separados por comas
        '-------------------------------------------------------------
        
        
        'Empresa
        '------------------------------------------------------------------------------------
        pos1 = pos2 + 2
        pos2 = InStr(pos1, Parametros, "@") - 1
        Empresa = Mid(Parametros, pos1, pos2 - pos1 + 1)
        
        Flog.writeline "Posicion 1 = " & pos1
        Flog.writeline "Pos 2 = " & pos2
        Flog.writeline "Parametro Empresa = " & Empresa
        Flog.writeline
        '------------------------------------------------------------------------------------
        
        'Fecha Desde
        '------------------------------------------------------------------------------------
        pos1 = pos2 + 2
        pos2 = InStr(pos1, Parametros, "@") - 1
        Fechadesde = Mid(Parametros, pos1, pos2 - pos1 + 1)
        
        Flog.writeline "Posicion 1 = " & pos1
        Flog.writeline "Pos 2 = " & pos2
        Flog.writeline "Parametro Fecha Desde = " & Fechadesde
        Flog.writeline
        '------------------------------------------------------------------------------------
        
        'Fecha Hasta
        '------------------------------------------------------------------------------------
        pos1 = pos2 + 2
        pos2 = InStr(pos1, Parametros, "@") - 1
        Fechahasta = Mid(Parametros, pos1, pos2 - pos1 + 1)
        
        Flog.writeline "Posicion 1 = " & pos1
        Flog.writeline "Pos 2 = " & pos2
        Flog.writeline "Parametro Fecha Hasta = " & Fechahasta
        Flog.writeline
        '------------------------------------------------------------------------------------
        
        'Tipo de Nomina
        '------------------------------------------------------------------------------------
        pos1 = pos2 + 2
        pos2 = Len(Parametros)
        tNomina = Mid(Parametros, pos1, pos2 - pos1 + 1)
        
        Flog.writeline "Posicion 1 = " & pos1
        Flog.writeline "Pos 2 = " & pos2
        Flog.writeline "Parametro Tipo de Nomina = " & tNomina
        Flog.writeline
        '------------------------------------------------------------------------------------
    End If
Else
    Flog.writeline "parametros nulos"
End If
Flog.writeline
Flog.writeline Espacios(Tabulador * 0) & "Terminó de levantar los parametros "
Flog.writeline


'Configuracion del Reporte

Flog.writeline "**************************columna 44"
StrSql = "SELECT * FROM confrep"
Select Case tNomina
    Case 1:
        StrSql = StrSql & " WHERE repnro = 186 "
    Case 2:
        StrSql = StrSql & " WHERE repnro = 256 "
    Case 3:
        StrSql = StrSql & " WHERE repnro = 257 "
    Case Else:
        StrSql = StrSql & " WHERE repnro = 186 "
End Select
StrSql = StrSql & " AND confrep.confnrocol = 44 "
'StrSql = StrSql & " ORDER BY confrep.confnrocol "
OpenRecordset StrSql, rs_Confrep
Flog.writeline StrSql
Do While Not rs_Confrep.EOF
 Flog.writeline "--->" & rs_Confrep!confnrocol & "||" & rs_Confrep!confval & "||" & rs_Confrep!confval2
 rs_Confrep.MoveNext
Loop
'rs_Confrep.Close

Flog.writeline "************************** Fin col 44"




StrSql = "SELECT * FROM confrep"
Select Case tNomina
    Case 1:
        StrSql = StrSql & " WHERE repnro = 186 "
    Case 2:
        StrSql = StrSql & " WHERE repnro = 256 "
    Case 3:
        StrSql = StrSql & " WHERE repnro = 257 "
    Case Else:
        StrSql = StrSql & " WHERE repnro = 186 "
End Select
StrSql = StrSql & " AND confrep.confnrocol not in (44,49)"
StrSql = StrSql & " ORDER BY confrep.confnrocol "
        

OpenRecordset StrSql, rs_Confrep
If rs_Confrep.EOF Then
    Select Case tNomina
        Case 1:
            Flog.writeline "No se encontró la configuración del Reporte 186"
        Case 2:
            Flog.writeline "No se encontró la configuración del Reporte 256"
        Case 3:
            Flog.writeline "No se encontró la configuración del Reporte 257"
        Case Else:
            Flog.writeline "No se encontró la configuración del Reporte 186"
    End Select
    
    Exit Sub
End If

'Flog.writeline "MyBeginTrans"

'Comienzo la transaccion
'MyBeginTrans

UltimoEmpleado = -1
Num_linea = 1


    StrSql = "SELECT * FROM proceso "
    StrSql = StrSql & " INNER JOIN cabliq ON cabliq.pronro = proceso.pronro "
    StrSql = StrSql & " INNER JOIN  tipoproc ON proceso.tprocnro = tipoproc.tprocnro" 'para sacar el ajugcias
    StrSql = StrSql & " INNER JOIN empleado ON cabliq.empleado = empleado.ternro "
    StrSql = StrSql & " Where proceso.pronro IN (" & Lista_Pro & ")"
    StrSql = StrSql & " ORDER BY empleado.ternro, proceso.pronro"
    OpenRecordset StrSql, rs_Empleados
    
    If rs_Empleados.State = adStateOpen Then
        Flog.writeline "busco los empleados"
    Else
        Flog.writeline "se supero el tiempo de espera "
        HuboError = True
    End If
    
If Not HuboError Then
    
        'seteo de las variables de progreso
        Progreso = 0
          
          'Obtengo la cantidad real de empleados
            StrSql = "SELECT distinct (empleado) FROM proceso  "
            StrSql = StrSql & " INNER JOIN cabliq ON cabliq.pronro = proceso.pronro  "
            StrSql = StrSql & " INNER JOIN empleado ON cabliq.empleado = empleado.ternro "
            StrSql = StrSql & " Where proceso.pronro IN (" & Lista_Pro & ")"
            OpenRecordset StrSql, rs_CantEmpleados
            
          CEmpleadosAProc = rs_CantEmpleados.RecordCount
        If CEmpleadosAProc = 0 Then
           Flog.writeline "no hay empleados"
           CEmpleadosAProc = 1
        End If
        IncPorc = (99 / CEmpleadosAProc)
        Flog.writeline
        Flog.writeline
        
        'Inicializo la cantidad de empleados con errores a 0
        CantEmplError = 0
        CantEmplSinError = 0
        
        'Busco el numero del tipo de estructura de la caja CCAF
        StrSql = "SELECT * FROM confrep "
        StrSql = StrSql & " WHERE confnrocol=83 AND conftipo='TE'"
        OpenRecordset StrSql, rs_Conceptos
        If Not rs_Conceptos.EOF Then
            teCajaCCAF = rs_Conceptos!confval
        Else
            teCajaCCAF = 0
        End If
        rs_Conceptos.Close
        
                
        'Busco el codigo de la estructura CAJA del cual depende la empresa.
        StrSql = "SELECT * FROM estruc_depende ed"
        StrSql = StrSql & " INNER JOIN empresa e ON e.estrnro = ed.estrnro1"
        StrSql = StrSql & " Where ed.tenro1 = 10 And e.Empnro = " & Empresa
        StrSql = StrSql & " and ed.tenro2 =" & teCajaCCAF
        OpenRecordset StrSql, rs_Conceptos
        If Not rs_Conceptos.EOF Then
            codCaja = rs_Conceptos!estrnro2
        Else
            codCaja = 0
        End If
        rs_Conceptos.Close
        '-----------------------------------------------------------------
        
    Do While Not rs_Empleados.EOF
       
        MyBeginTrans
          rs_Confrep.MoveFirst
          
          total_APVI = 0
          total_APVC = 0
          
          If rs_Empleados!ternro <> UltimoEmpleado Then   'Es el primero
                    
               UltimoEmpleado = rs_Empleados!ternro
                Flog.writeline "_______________________________________________________________________"
                
                'Buscar el apellido y nombre
                    StrSql = "SELECT * FROM tercero WHERE ternro = " & rs_Empleados!ternro
                    OpenRecordset StrSql, rs_Tercero
                    If Not rs_Tercero.EOF Then
                        Apellido = Left(rs_Tercero!terape, 30)
                        Apellido2 = Left(rs_Tercero!terape2, 30)
                        
                        LongNombreEmp1 = Len(rs_Tercero!ternom)
                        
                        If Not EsNulo(rs_Tercero!ternom2) Then
                            LongNombreEmp2 = Len(rs_Tercero!ternom2)
                        Else
                            LongNombreEmp2 = 0
                        End If
                        
                        If (((LongNombreEmp1 + LongNombreEmp2) <= 29) And (Not EsNulo(rs_Tercero!ternom2))) Then
                            NombreEmp = rs_Tercero!ternom & " " & rs_Tercero!ternom2
                        
                        Else
                            NombreEmp = rs_Tercero!ternom
                        
                        End If
                    Else
                        Flog.writeline Espacios(Tabulador * 1) & "ERROR al obtener Apellido o Nombre del Empleado"
                        Exit Sub
                    End If
                    Flog.writeline
                    Flog.writeline
                    
                    Flog.writeline "Empleado: ------------------->" & rs_Empleados!empleg & "  " & Apellido & "  " & NombreEmp
                    Flog.writeline
                    
                    'Inicializar Arreglos de totales
                    Flog.writeline "Inicializar Arreglos de totales"
                    
                    For Contador = 1 To 210 'antes 110
'                        If (Contador = 78) Then 'HJI Existen para esta columna valores validos con cero.
'                           arreglo(Contador) = -1 'Todo el arreglo deberia ser inicializado en -1 HJI
'                        Else
                           arreglo(Contador) = 0
                        'End If
                        
                        arregloEstruc(Contador) = ""
                    Next Contador
                    
                    For Contador = 1 To 30
                        arregloMov(Contador) = 0
                        arregloFecD(Contador) = vbNull
                        arregloFecH(Contador) = vbNull
                        
'                        arregloAPVI(Contador).Cod = 0
'                        arregloAPVI(Contador).Contrato = ""
'                        arregloAPVI(Contador).Cotiza = 0
'                        arregloAPVI(Contador).Depositos = 0
'                        arregloAPVI(Contador).FPago = 0
'
'                        arregloAPVC(Contador).Cod = 0
'                        arregloAPVC(Contador).Contrato = ""
'                        arregloAPVC(Contador).Cotiza = 0
'                        arregloAPVC(Contador).Depositos = 0
'                        arregloAPVC(Contador).FPago = 0
                    Next Contador

                    'Flog.writeline "Inicializar Arreglos de totales"
          End If
                
                Do While Not rs_Confrep.EOF
                    
                    Flog.writeline "Columna " & rs_Confrep!confnrocol
                    Select Case UCase(rs_Confrep!conftipo)
                    Case "AC":
                        Flog.writeline Espacios(Tabulador * 1) & "AC Busca Acumulador: " & rs_Confrep!confval
                        StrSql = "SELECT * FROM acu_liq WHERE cliqnro = " & rs_Empleados!cliqnro & _
                                 " AND acunro =" & rs_Confrep!confval
                        OpenRecordset StrSql, rs_Acu_liq
                        
                        If Not rs_Acu_liq.EOF Then
                                arreglo(rs_Confrep!confnrocol) = arreglo(rs_Confrep!confnrocol) + Abs(rs_Acu_liq!almonto)
                                Flog.writeline Espacios(Tabulador * 2) & "Encontro " & Abs(rs_Acu_liq!almonto) & " Acumulado " & arreglo(rs_Confrep!confnrocol)
                        Else
                                Flog.writeline Espacios(Tabulador * 2) & "No Encontro"
                        End If
                    Case "CO":
                    
                        Flog.writeline Espacios(Tabulador * 1) & "CO Busca Concepto: " & rs_Confrep!confval2
                        StrSql = "SELECT * FROM concepto "
                        StrSql = StrSql & " WHERE (concepto.conccod = " & rs_Confrep!confval
                        StrSql = StrSql & " OR concepto.conccod = '" & rs_Confrep!confval2 & "')"
                        OpenRecordset StrSql, rs_Conceptos
                        If Not rs_Conceptos.EOF Then
                            StrSql = "SELECT * FROM detliq WHERE concnro = " & rs_Conceptos!concnro & _
                                     " AND cliqnro =" & rs_Empleados!cliqnro
                            OpenRecordset StrSql, rs_Detliq
                            If Not rs_Detliq.EOF Then
                                If rs_Detliq!dlimonto <> 0 Then
                                    If (rs_Confrep!confnrocol > 84 And rs_Confrep!confnrocol < 90) Then 'And rs_Confrep!confnrocol = 87 And rs_Confrep!confnrocol = 90 Then
                                        If codCaja = 0 Then
                                            arreglo(rs_Confrep!confnrocol) = arreglo(rs_Confrep!confnrocol) + Abs(rs_Detliq!dlimonto)
                                        Else
                                            'nuevo seba 26/05/2015
                                            If rs_Confrep!confval <> 0 Then
                                                If codCaja = rs_Confrep!confval Then
                                                    arreglo(rs_Confrep!confnrocol) = arreglo(rs_Confrep!confnrocol) + Abs(rs_Detliq!dlimonto)
                                                End If
                                            Else
                                                arreglo(rs_Confrep!confnrocol) = arreglo(rs_Confrep!confnrocol) + Abs(rs_Detliq!dlimonto)
                                            End If
                                        End If
                                    Else
                                        arreglo(rs_Confrep!confnrocol) = arreglo(rs_Confrep!confnrocol) + Abs(rs_Detliq!dlimonto)
                                    End If
                                End If
                                Flog.writeline Espacios(Tabulador * 2) & "Encontro OK " & Abs(rs_Detliq!dlimonto) & " Acum " & arreglo(rs_Confrep!confnrocol)
                            End If
                         Else
                            Flog.writeline Espacios(Tabulador * 2) & "No Encontro"
                        End If
                    
                    Case "PCO":
                        Flog.writeline Espacios(Tabulador * 1) & "PCO Busca Parametro de Concepto: " & rs_Confrep!confval2
                        StrSql = "SELECT * FROM concepto "
                        StrSql = StrSql & " WHERE (concepto.conccod = " & rs_Confrep!confval
                        StrSql = StrSql & " OR concepto.conccod = '" & rs_Confrep!confval2 & "')"
                        OpenRecordset StrSql, rs_Conceptos
                        If Not rs_Conceptos.EOF Then
                            StrSql = "SELECT * FROM detliq WHERE concnro = " & rs_Conceptos!concnro & _
                                     " AND cliqnro =" & rs_Empleados!cliqnro
                            OpenRecordset StrSql, rs_Detliq
                            If Not rs_Detliq.EOF Then
                                If rs_Detliq!dlicant <> 0 Then
                                    arreglo(rs_Confrep!confnrocol) = arreglo(rs_Confrep!confnrocol) + Abs(rs_Detliq!dlicant)
                                End If
                                Flog.writeline Espacios(Tabulador * 2) & "Encontro OK " & Abs(rs_Detliq!dlicant) & " Acum " & arreglo(rs_Confrep!confnrocol)
                            End If
                        Else
                            Flog.writeline Espacios(Tabulador * 2) & "No Encontro"
                        End If
                    
                    Case "APV":
                        'Flog.writeline "-------------------- DEBUG!!!"
                        Flog.writeline Espacios(Tabulador * 1) & "APV Buscando Concepto: " & rs_Confrep!confval2
                        StrSql = "SELECT * FROM concepto "
                        StrSql = StrSql & " WHERE (concepto.conccod = " & rs_Confrep!confval
                        StrSql = StrSql & " OR concepto.conccod = '" & rs_Confrep!confval2 & "')"
                        OpenRecordset StrSql, rs_Conceptos
                        If Not rs_Conceptos.EOF Then
                            StrSql = "SELECT * FROM detliq WHERE concnro = " & rs_Conceptos!concnro & _
                                     " AND cliqnro =" & rs_Empleados!cliqnro
                            OpenRecordset StrSql, rs_Detliq
                            If Not rs_Detliq.EOF Then
                                If rs_Detliq!dlicant <> 0 Then
                                    'arreglo(rs_Confrep!confnrocol) = Abs(rs_Detliq!dlicant)
                                    'arreglo(rs_Confrep!confnrocol + 3) = Abs(rs_Detliq!dlimonto)
                                    Select Case CLng(rs_Confrep!confnrocol)
                                        Case 40:
                                            Flog.writeline Espacios(Tabulador * 2) & "Encuentra institucion APVI"
                                            'Encontre una institucion liquidada
                                            'total_APVI = total_APVI + 1
                                            'arregloAPVI(total_APVI).Cod = Abs(rs_Detliq!dlicant)
                                            'arregloAPVI(total_APVI).Cotiza = Abs(rs_Detliq!dlimonto)   'POSIBLE COLUMNA 43 mdf
                                            
                                            'Busco OTRO concepto liquidado asociado a la institucuion (mismo parametro)
                                            StrSql = "SELECT detliq.dlimonto,confrep.confval2"
                                            StrSql = StrSql & " From detliq"
                                            StrSql = StrSql & " INNER JOIN concepto ON concepto.concnro = detliq.concnro"
                                            StrSql = StrSql & " INNER JOIN confrep ON confrep.confval2 = concepto.conccod"
                                            StrSql = StrSql & " AND confrep.repnro = " & rs_Confrep!repnro
                                            StrSql = StrSql & " AND confrep.confnrocol = 44"
                                            StrSql = StrSql & " and confrep.confval2 = " & rs_Confrep!confval2
                                            StrSql = StrSql & " WHERE cliqnro = " & rs_Empleados!cliqnro
                                            StrSql = StrSql & " AND detliq.dlicant = " & rs_Detliq!dlicant
                                            StrSql = StrSql & " and confrep.confval2 = " & rs_Confrep!confval2 'mdf
                                            OpenRecordset StrSql, rs_Aux
                                            
                                            Flog.writeline StrSql & "----" & rs_Aux.RecordCount
                                            If Not rs_Aux.EOF Then
                                              Flog.writeline "IF, columna 44"
                                              Flog.writeline "Entre por el concepto " & rs_Aux!confval2
                                              Do While Not rs_Aux.EOF 'mdf
                                                total_APVI = total_APVI + 1
                                                Flog.writeline "Indice incrementado--->" & total_APVI
                                                Flog.writeline Espacios(Tabulador * 3) & "Encuentra Cotizacion Dep Convenidos (44)"
                                                arregloAPVI(total_APVI).Cod = Abs(rs_Detliq!dlicant)
                                                arregloAPVI(total_APVI).Depositos = Abs(rs_Aux!dlimonto)
                                                arregloAPVI(total_APVI).Cotiza = 0
                                                Flog.writeline "Deposito: " & arregloAPVI(total_APVI).Depositos
                                                rs_Aux.MoveNext '----mdf
                                              Loop
                                            Else
                                              Flog.writeline "ELSE, columna 43"
                                              total_APVI = total_APVI + 1
                                              Flog.writeline "Indice INCREMENTADO --->" & total_APVI
                                              arregloAPVI(total_APVI).Cod = Abs(rs_Detliq!dlicant)
                                              arregloAPVI(total_APVI).Cotiza = Abs(rs_Detliq!dlimonto)   'POSIBLE COLUMNA 43 mdf
                                                'El registro no esta completo
                                              Flog.writeline Espacios(Tabulador * 3) & "NO encuentra Cotizacion Dep Convenidos (44)"
                                                'total_APVI = total_APVI - 1
                                                'HuboError = True
                                                arregloAPVI(total_APVI).Depositos = 0
                                            End If
                                       ' Flog.writeline "-------------------- FIN DEBUG!!"
                                        Case 45:
                                            'Busco OTRO concepto liquidado asociado a la institucuion (mismo parametro)
                                            Flog.writeline Espacios(Tabulador * 2) & "Encuentra institucion APVC"
                                            total_APVC = total_APVC + 1
                                            arregloAPVI(total_APVC).Cod = Abs(rs_Detliq!dlicant)
                                            arregloAPVI(total_APVC).Cotiza = Abs(rs_Detliq!dlimonto)
                                            
                                            'Busco OTRO concepto liquidado asociado a la institucuion (mismo parametro)
                                            StrSql = "SELECT detliq.dlimonto"
                                            StrSql = StrSql & " From detliq"
                                            StrSql = StrSql & " INNER JOIN concepto ON concepto.concnro = detliq.concnro"
                                            StrSql = StrSql & " INNER JOIN confrep ON confrep.confval2 = concepto.conccod"
                                            StrSql = StrSql & " AND confrep.repnro = " & rs_Confrep!repnro
                                            StrSql = StrSql & " AND confrep.confnrocol = 49"
                                            StrSql = StrSql & " WHERE cliqnro = " & rs_Empleados!cliqnro
                                            StrSql = StrSql & " AND detliq.dlicant = " & rs_Detliq!dlicant
                                            OpenRecordset StrSql, rs_Aux
                                            If Not rs_Aux.EOF Then
                                                Flog.writeline Espacios(Tabulador * 3) & "Encuentra Cotizacion Empl"
                                                arregloAPVC(total_APVC).Depositos = Abs(rs_Aux!dlimonto)
                                            Else
                                                'El registro no esta completo
                                                Flog.writeline Espacios(Tabulador * 3) & "NO encuentra Cotizacion Empl"
                                                'total_APVC = total_APVC - 1
                                                'HuboError = True
                                                arregloAPVC(total_APVC).Depositos = 0
                                            End If
                                    End Select
                                End If
                            End If
                        End If
                        
                    Case "TE": 'tipo estructura
                            Flog.writeline Espacios(Tabulador * 1) & "TE Buscando Estructura Tipo: " & rs_Confrep!confval
                            StrSql = " SELECT estructura.estrnro, estructura.estrdabr, estructura.estrcodext, htetdesde, htethasta "
                            StrSql = StrSql & " FROM his_estructura "
                            StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro "
                            StrSql = StrSql & " WHERE ternro = " & rs_Empleados!ternro & " AND "
                            StrSql = StrSql & " his_estructura.tenro = " & rs_Confrep!confval & " And "
                            StrSql = StrSql & " (his_estructura.htetdesde <= " & ConvFecha(Fechahasta) & ") And "
                            StrSql = StrSql & " ((" & ConvFecha(Fechahasta) & " <= his_estructura.htethasta) or (his_estructura.htethasta is null))"
                            OpenRecordset StrSql, rs_Estructura
                            If Not rs_Estructura.EOF Then
                                Flog.writeline Espacios(Tabulador * 2) & "Estructura: " & rs_Estructura!estrnro & " - " & rs_Estructura!estrdabr
                                If rs_Confrep!confnrocol = 37 Then
                                    arregloEstruc(37) = rs_Estructura!estrdabr
                                    Flog.writeline Espacios(Tabulador * 2) & "Encontro OK "
                                Else
                                    StrSql = "SELECT * FROM estr_cod WHERE estrnro =" & rs_Estructura!estrnro
                                    StrSql = StrSql & " AND tcodnro = 38"
                                    OpenRecordset StrSql, rs_Estr_cod
                                    If Not rs_Estr_cod.EOF Then
                                        arregloEstruc(rs_Confrep!confnrocol) = IIf(EsNulo(rs_Estr_cod!nrocod), "", CStr(rs_Estr_cod!nrocod))
                                        Flog.writeline Espacios(Tabulador * 2) & "Encontro OK "
                                        Flog.writeline
                                    Else
                                        Flog.writeline Espacios(Tabulador * 2) & "  No se encontró el codigo Interno para el Tipo de Estructura."
                                        Flog.writeline
                                    End If
                                    rs_Estr_cod.Close
                                End If
                            Else
                                Flog.writeline Espacios(Tabulador * 2) & "Tipo de Estructura: " & rs_Confrep!confval
                                Flog.writeline Espacios(Tabulador * 2) & "No se encontró la estructura para ese empleado en la fecha del filtro ."
                                Flog.writeline
                            End If
                            rs_Estructura.Close
                            
                    Case "TM": 'tipo movimiento
                            Flog.writeline Espacios(Tabulador * 1) & "TM Buscando Tipo Mov "
                            'Hacer case de tipo de Movimiento y generar el array con las fechas correspondientes
                            
                            'If (rs_Empleados!ternro <> UltimoEmpleado) Then
                                total_mov = 0
                            'End If
                            'ALTA
                            Flog.writeline Espacios(Tabulador * 2) & "TM Alta busca Fases "
                            StrSql = "SELECT fases.altfec, fases.bajfec FROM fases "
                            StrSql = StrSql & " WHERE fases.real = -1 "
                            StrSql = StrSql & " AND fases.altfec >=" & ConvFecha(Fechadesde)
                            StrSql = StrSql & " AND fases.altfec <= " & ConvFecha(Fechahasta)
                            StrSql = StrSql & " AND empleado = " & rs_Empleados!ternro
                            OpenRecordset StrSql, rs_Fases
                            Do While Not rs_Fases.EOF
                                total_mov = total_mov + 1
                                arregloMov(total_mov) = 1   'fijo
                                If Not EsNulo(rs_Fases!altfec) Then
                                    arregloFecD(total_mov) = rs_Fases!altfec
                                    arregloFecH(total_mov) = rs_Fases!altfec
                                End If
                                rs_Fases.MoveNext
                            Loop
                            
                            'BAJA"
                            Flog.writeline Espacios(Tabulador * 2) & "TM Baja busca Fases "
                            StrSql = "SELECT fases.caunro, fases.altfec, fases.bajfec FROM fases "
                            StrSql = StrSql & " WHERE fases.real = -1 "
                            StrSql = StrSql & " AND fases.bajfec >=" & ConvFecha(Fechadesde)
                            StrSql = StrSql & " AND fases.bajfec <= " & ConvFecha(Fechahasta)
                            StrSql = StrSql & " AND fases.empleado = " & rs_Empleados!ternro
                            OpenRecordset StrSql, rs_Fases
                            Do While Not rs_Fases.EOF
                            
                                total_mov = total_mov + 1 'esto estaba 27/01/12
                                'segun la causa ==> busco la estructura y el codigo asociado
                                StrSql = "SELECT * FROM estr_cod "
                                StrSql = StrSql & " INNER JOIN causa_sitrev ON causa_sitrev.estrnro = estr_cod.estrnro"
                                StrSql = StrSql & " WHERE tcodnro = 38"
                                StrSql = StrSql & " AND causa_sitrev.caunro = " & rs_Fases!caunro
                                OpenRecordset StrSql, rs_Estr_cod
                                Flog.writeline "consulta" & StrSql
                                If Not rs_Estr_cod.EOF Then

                                    arregloMov(total_mov) = IIf(EsNulo(rs_Estr_cod!nrocod), "", Left(CStr(rs_Estr_cod!nrocod), 6))
                                    Flog.writeline Espacios(Tabulador * 2) & " OK "
                                    Flog.writeline
                                Else
                                    Flog.writeline Espacios(Tabulador * 2) & "  No se encontró el codigo Interno para el Movimiento."
                                    Flog.writeline
                                End If
                                
                                If Not EsNulo(rs_Fases!bajfec) Then
                                    arregloFecD(total_mov) = rs_Fases!bajfec
                                    arregloFecH(total_mov) = rs_Fases!bajfec
                                End If
                                
                                rs_Fases.MoveNext
                            Loop
                            

                            
                            'Licencias
                            Flog.writeline Espacios(Tabulador * 2) & "TM Licencias "
                            StrSql = " SELECT emp_lic.elfechadesde, emp_lic.elfechahasta, emp_lic.tdnro,emp_lic.emp_licnro FROM emp_lic "
                            StrSql = StrSql & " WHERE emp_lic.empleado= " & rs_Empleados!ternro
                            StrSql = StrSql & " AND ((emp_lic.elfechadesde <= " & ConvFecha(Fechadesde)
                            StrSql = StrSql & " AND emp_lic.elfechahasta >= " & ConvFecha(Fechadesde) & ")"
                            StrSql = StrSql & " OR (emp_lic.elfechadesde >=" & ConvFecha(Fechadesde) & " AND emp_lic.elfechahasta <=" & ConvFecha(Fechahasta) & ")" ' )"
                            StrSql = StrSql & " OR (emp_lic.elfechadesde >=" & ConvFecha(Fechadesde) & " and emp_lic.elfechadesde <=" & ConvFecha(Fechahasta) & " AND emp_lic.elfechahasta >=" & ConvFecha(Fechahasta) & "))" 'MDF
                            OpenRecordset StrSql, rs_Aux
                            Flog.writeline "consulta licencia" & StrSql
                            Do While Not rs_Aux.EOF
                                total_mov = total_mov + 1
                                Flog.writeline Espacios(Tabulador * 2) & " Licencia Encontrada: " & rs_Aux!emp_licnro
                                StrSql = "SELECT * FROM estr_cod "
                                StrSql = StrSql & " INNER JOIN csijp_srtd ON estr_cod.estrnro = csijp_srtd.estrnro "
                                StrSql = StrSql & " WHERE tcodnro = 38"
                                StrSql = StrSql & " AND csijp_srtd.tdnro = " & rs_Aux!tdnro
                                OpenRecordset StrSql, rs_Estr_cod
                                If Not rs_Estr_cod.EOF Then
                                                                
                                   arregloMov(total_mov) = IIf(EsNulo(rs_Estr_cod!nrocod), "", Left(CStr(rs_Estr_cod!nrocod), 6))
                                   Flog.writeline Espacios(Tabulador * 2) & " OK tipo codigo"
                                   Flog.writeline
                                Else
                                   Flog.writeline Espacios(Tabulador * 2) & "  No se encontró el codigo Interno para el Movimiento."
                                   Flog.writeline
                                End If
                                
                                If CDate(rs_Aux!elfechadesde) < CDate(Fechadesde) Then
                                    arregloFecD(total_mov) = Fechadesde
                                Else
                                    arregloFecD(total_mov) = rs_Aux!elfechadesde
                                End If
                               If CDate(rs_Aux!elfechahasta) > CDate(Fechahasta) Then
                                    arregloFecH(total_mov) = Fechahasta
                               Else
                                    arregloFecH(total_mov) = rs_Aux!elfechahasta
                                End If
                            
                                rs_Aux.MoveNext
                            Loop
                            
                            'me fijo si el proceso esta marcado como post pago
                            'recorro la lista de procesos y busco si alguno es finiquito
                            procesos = Split(Lista_Pro, ",")
                            For I = 0 To UBound(procesos)
                                StrSql = "SELECT * FROM proceso "
                                StrSql = StrSql & " INNER JOIN cabliq ON cabliq.pronro = proceso.pronro "
                                StrSql = StrSql & " INNER JOIN  tipoproc ON proceso.tprocnro = tipoproc.tprocnro" 'para sacar el ajugcias
                                StrSql = StrSql & " INNER JOIN empleado ON cabliq.empleado = empleado.ternro "
                                StrSql = StrSql & " Where proceso.pronro = (" & procesos(I) & ")"
                                StrSql = StrSql & " ORDER BY empleado.ternro, proceso.pronro"
                                OpenRecordset StrSql, rs_Empleados1
                                If Not rs_Empleados1.EOF Then
                                    If rs_Empleados1!postfiniquito = -1 Then
                                        total_mov = total_mov + 1
                                        arregloMov(total_mov) = 12
                                        arregloFecD(total_mov) = rs_Empleados!profecini
                                        arregloFecH(total_mov) = rs_Empleados!profecfin
                                    End If
                                End If
                            Next
                            'hasta aca
                            
                            'FGZ 26/01/2012 --------------------------------------------
                            'codigo seba 27/01/2012
                            'busca el concepto puente que me indica si hubo cambio de contrato
                            'Flog.writeline Espacios(Tabulador * 1) & "busca el concepto puente que me indica si hubo cambio de contrato"
                            'Flog.writeline Espacios(Tabulador * 1) & "CO Busca Concepto de cambio de contrato: " & rs_Confrep!confval2
                            'StrSql = "SELECT * FROM concepto "
                            'StrSql = StrSql & " WHERE (concepto.conccod = " & rs_Confrep!confval
                            'StrSql = StrSql & " OR concepto.conccod = '" & rs_Confrep!confval2 & "')"
                            'OpenRecordset StrSql, rs_Conceptos
                            'If Not rs_Conceptos.EOF Then
                            '    StrSql = "SELECT * FROM detliq WHERE concnro = " & rs_Conceptos!concnro & _
                            '             " AND cliqnro =" & rs_Empleados!cliqnro
                            '    OpenRecordset StrSql, rs_Detliq
                            '    If Not rs_Detliq.EOF Then
                            '        If rs_Detliq!dlimonto <> 0 Then
                            
                            '            cambioContrato = True
                            '            Flog.writeline "Hay cambio de contrato"
                            '        Else
                            '            cambioContrato = False
                            '            Flog.writeline "No hay cambio de contrato"
                            '        End If
                            '        Flog.writeline Espacios(Tabulador * 2) & "Encontro OK " & Abs(rs_Detliq!dlimonto) & " Acum " & arreglo(rs_Confrep!confnrocol)
                            '    End If
                            ' Else
                            '    Flog.writeline Espacios(Tabulador * 2) & "No Encontro"
                            'End If
                            
                            'If cambioContrato Then
                            '    If ((Not EsNulo(arreglo(100))) And (Not EsNulo(arreglo(200))) And (arreglo(100) <> 0) And (arreglo(200) <> 0)) Then
                            '        total_mov = total_mov + 1
                            
                            '        Flog.writeline "abre apertura porque  son distinto de ceros"
                            '    Else
                            '       Flog.writeline "no abre apertura"
                            '    End If
        
                            'Else
                            '    Flog.writeline "no hay cambio de contrato"
                            
                            'End If
                           'hasta aca
                            
                           'FGZ 26/01/2012 --------------------------------------------
                            
                            
                            
                            If total_mov = 0 Then
                                Flog.writeline Espacios(Tabulador * 2) & "Tipo de Estructura: " & rs_Confrep!confval
                                Flog.writeline Espacios(Tabulador * 2) & "No se encontraron movimientos."
                                Flog.writeline
                            Else
                                Flog.writeline Espacios(Tabulador * 2) & "se encontraron " & total_mov & " movimientos."
                                Flog.writeline
                            End If
                            
                    Case "CTE": 'Constante
                                Flog.writeline Espacios(Tabulador * 1) & "CTE Buscando Constantes "
                                If rs_Confrep!confval2 = "" Or EsNulo(rs_Confrep!confval2) Then
                                    'Numerica
                                    arreglo(rs_Confrep!confnrocol) = rs_Confrep!confval
                                Else
                                    'Alfanumerica
                                    arregloEstruc(rs_Confrep!confnrocol) = rs_Confrep!confval2
                                End If
                                
                    Case "TN": 'tipo nuevo
                            'FGZ 26/01/2012 --------------------------------------------
                            'codigo seba 27/01/2012
                            'busca el concepto puente que me indica si hubo cambio de contrato
                            If total_mov = 0 Then
                                total_mov = 1
                            End If
                            Flog.writeline Espacios(Tabulador * 1) & "busca el concepto puente que me indica si hubo cambio de contrato"
                            Flog.writeline Espacios(Tabulador * 1) & "CO Busca Concepto de cambio de contrato: " & rs_Confrep!confval2
                            StrSql = "SELECT * FROM concepto "
                            StrSql = StrSql & " WHERE (concepto.conccod = " & rs_Confrep!confval
                            StrSql = StrSql & " OR concepto.conccod = '" & rs_Confrep!confval2 & "')"
                            OpenRecordset StrSql, rs_Conceptos
                            If Not rs_Conceptos.EOF Then
                                StrSql = "SELECT * FROM detliq WHERE concnro = " & rs_Conceptos!concnro & _
                                         " AND cliqnro =" & rs_Empleados!cliqnro
                                OpenRecordset StrSql, rs_Detliq
                                Flog.writeline "Busca el concepto liquidado" & StrSql
                                If Not rs_Detliq.EOF Then
                                    If rs_Detliq!dlimonto <> 0 Then
                                        'arreglo(rs_Confrep!confnrocol) = arreglo(rs_Confrep!confnrocol) + Abs(rs_Detliq!dlimonto)
                                        cambioContrato = True
                                        estructuraContrato = CInt(rs_Detliq!dlicant) 'uso cint porque siempre es 861 y no rompe
                                        Flog.writeline "Hay cambio de contrato"
                                    Else
                                        cambioContrato = False
                                        Flog.writeline "No hay cambio de contrato"
                                    End If
                                    Flog.writeline Espacios(Tabulador * 2) & "Encontro OK " & Abs(rs_Detliq!dlimonto) & " Acum " & arreglo(rs_Confrep!confnrocol)
                                Else
                                    cambioContrato = False
                                End If
                             Else
                                Flog.writeline Espacios(Tabulador * 2) & "No Encontro"
                            End If
                            
                            If cambioContrato Then
                                If ((Not EsNulo(arreglo(100))) And (Not EsNulo(arreglo(200))) And (arreglo(100) <> 0) And (arreglo(200) <> 0)) Then
                                    total_mov = total_mov + 1
                                    arregloMov(total_mov) = 8
                                    'arregloMov(total_mov) = IIf(EsNulo(rs_Estr_cod!nrocod), "", Left(CStr(rs_Estr_cod!nrocod), 6))
                                    Flog.writeline "abre apertura porque  son distinto de ceros"
                                    'nuevo 01022012
                                    sql = "SELECT * FROM his_estructura where estrnro=" & estructuraContrato & " AND ternro=" & rs_Empleados!ternro
                                    Flog.writeline "busca las fechas del cambio de contrato" & sql
                                    OpenRecordset sql, rs_Contratos
                                    If Not rs_Contratos.EOF Then
                                        If Not EsNulo(rs_Contratos!htetdesde) Then
                                            dia = Day(rs_Contratos!htetdesde)
                                            dia = dia - 1
                                            'mes = Month(rs_Contratos!htetdesde)
                                            'Anio = Year(rs_Contratos!htetdesde)
                                            arregloFecD(1) = rs_Contratos!htetdesde - dia
                                        Else
                                            arregloFecD(1) = vbNull
                                        End If
                                        If Not EsNulo(rs_Contratos!htethasta) Then
                                            arregloFecH(1) = rs_Contratos!htethasta
                                            arregloFecD(total_mov) = DateAdd("y", 1, rs_Contratos!htethasta)
                                        Else
                                            arregloFecH(1) = vbNull
                                            arregloFecD(total_mov) = vbNull
                                        End If
                                        arregloFecH(total_mov) = vbNull
                                        
                                        Flog.writeline "Fechas encontradas=" & arregloFecH(total_mov) & arregloFecD(total_mov)
                                    End If
                                    'hasta aca
                                Else
                                   Flog.writeline "no abre apertura"
                                End If
        
                            Else
                                Flog.writeline "no hay cambio de contrato"
                                'total_mov = total_mov + 1
                            End If
                           'hasta aca
                    
                    
                    
                    
                    
                    
                    
                    Case Else
                    
                    End Select
                    
                    rs_Confrep.MoveNext
                Loop
                

                
                'Reviso si es el ultimo empleado
                If EsElUltimoEmpleado(rs_Empleados, UltimoEmpleado) Then
                    'Inicializo
                        HuboError = False 'Para cada empleado
                        Errores = False 'En el proceso
                        EsFonasa = False
                        
                    '-----------------------------------------------------------------------------------
                    'Lleno con ceros las posiciones vacias del arreglo para que NO tire error al insertar
                    For Contador = 1 To 210 'antes 110
                        If arregloEstruc(Contador) = "" Then
                            arregloEstruc(Contador) = "0"
                        End If
                    Next
                    arregloEstruc(18) = "-1"
                    '-----------------------------------------------------------------------------------
                    ' Bloque Datos del Trabajador
                    ' ----------------------------------------------------------------
                    ' Buscar el Rut DEL EMPLEADO
                    Flog.writeline
                    Flog.writeline "----------------------------------------------------"
                    Flog.writeline "Procesando Campo 1 y 2. RUT y DV:  "
                    '28/10/2014
                    'StrSql = " SELECT nrodoc FROM tercero " & _
                    '         " INNER JOIN ter_doc  ON (tercero.ternro = ter_doc.ternro AND ter_doc.tidnro = 1) " & _
                    '         " WHERE tercero.ternro= " & rs_Empleados!ternro
                    
                    'Inicio
                    StrSql = " SELECT nrodoc FROM tercero "
                    StrSql = StrSql & " INNER JOIN ter_doc ON ter_doc.ternro =tercero.ternro "
                    StrSql = StrSql & " INNER JOIN tipodocu ON ter_doc.tidnro = tipodocu.tidnro "
                    StrSql = StrSql & " INNER JOIN tipodocu_pais on tipodocu_pais.tidnro = ter_doc.tidnro "
                    StrSql = StrSql & " AND tipodocu_pais.paisnro = 8"
                    StrSql = StrSql & " AND tipodocu_pais.tidcod = 1"
                    StrSql = StrSql & " WHERE Tercero.ternro = " & rs_Empleados!ternro
                    'fin
                    OpenRecordset StrSql, rs_Rut
                    
                    If Not rs_Rut.EOF Then
                        RUT = Mid(rs_Rut!nrodoc, 1, Len(rs_Rut!nrodoc) - 1)
                        RUT = Replace(RUT, "-", "")
                        DV = Right(rs_Rut!nrodoc, 1)
                        'HACER VALIDACION DE RUT Y DV
                        Flog.writeline Espacios(Tabulador * 1) & "RUT y DV Validos " & rs_Rut!nrodoc
                    Else
                        Flog.writeline Espacios(Tabulador * 1) & "Error al obtener los datos del RUT"
                        RUT = ""
                        DV = ""
                        HuboError = True
                    End If
                    Flog.writeline
                    Flog.writeline "Procesando Campo 3,4 y 5"
                    Flog.writeline Espacios(Tabulador * 1) & "Datos Obtenidos"
                    Flog.writeline
                    
                  '-----------------------------------------------------------------
                    'SEXO
                    Flog.writeline "Procesando Campo 6: Sexo "
                    If Not rs_Tercero.EOF Then
                        If rs_Tercero!tersex = -1 Then
                            Sexo = "M"
                        Else
                            Sexo = "F"
                        End If
                        Flog.writeline Espacios(Tabulador * 1) & "Sexo Obtenido "
                    Else
                        Flog.writeline Espacios(Tabulador * 1) & "ERROR No se encontró el Sexo del empleado"
                        Sexo = ""
                        HuboError = True
                    End If
                    Flog.writeline
                  '-----------------------------------------------------------------
                    
                  '----------------------------------------------------------------
                    'Busco la Nacionalidad
                    Flog.writeline "Procesando Campo 7: Nacionalidad "
                    StrSql = "SELECT * FROM nacionalidad"
                    StrSql = StrSql & " WHERE nacionalnro = " & rs_Tercero!nacionalnro
                    OpenRecordset StrSql, rs_Nacionalidad
                    If rs_Nacionalidad!nacionaldefault = -1 Then
                       Nacionalidad = 0
                    Else
                       Nacionalidad = 1
                    End If
                    Flog.writeline Espacios(Tabulador * 1) & " Nacionalidad Obtenida "
                    Flog.writeline
                  '----------------------------------------------------------------
                    
                  '----------------------------------------------------------------
                    'Busco el valor de Tipo Pago
                    Flog.writeline "Procesando Campo 8: Tipo Pago "
'                        If rs_Empleados!ajugcias = -1 Then
'                            TipoPago = 2
'                        Else
'                            TipoPago = 1
'                        End If
                         TipoPago = tNomina
                    Flog.writeline Espacios(Tabulador * 1) & " Tipo Pago Obtenido "
                    Flog.writeline
                  '----------------------------------------------------------------
                    
                  
                    
                    'Periodo de Remuneraciones   DESDE  Formato MMAAAA
                    'Este fue pasado por parametro
                    Flog.writeline "Procesando Campo 9: Periodo Remuneraciones Desde"
                    Flog.writeline Espacios(Tabulador * 1) & "Periodo Remuneraciones Desde Obtenido "
                    Flog.writeline
                    
                    'Periodo de Remuneraciones  HASTA   Formato MMAAAA
                    'Este fue pasado por parametro
                    Flog.writeline "Procesando Campo 10: Periodo Remuneraciones Hasta"
                    Flog.writeline Espacios(Tabulador * 1) & "Periodo Remuneraciones Hasta Obtenido "
                    Flog.writeline
                    
                    'Regimen Previsional
                    Flog.writeline "Procesando Campo 11: Regimen Previsional"
                    If arregloEstruc(11) = "0" Then
                        Flog.writeline Espacios(Tabulador * 1) & "ERROR No se encontro el codigo para el  Regimen Previsional"
                        HuboError = True
                    Else
                        Flog.writeline Espacios(Tabulador * 1) & "Regimen Previsional Obtenido "
                    End If
                    Flog.writeline
                    
                    'Tipo Trabajador
                    Flog.writeline "Procesando Campo 12: Tipo Trabajador"
                    'PERMITE 0   08-03-07
                   ' If arregloEstruc(12) = "0" Then
                    '    Flog.writeline Espacios(Tabulador * 1) & "ERROR No se encontro el codigo para Tipo de Trabajador"
                    '    Exit Sub
                    'End If
                    If IsNumeric(arregloEstruc(12)) = True Then
                            Flog.writeline Espacios(Tabulador * 1) & "Tipo de Trabajador Obtenido "
                    Else
                            Flog.writeline Espacios(Tabulador * 1) & "ERROR. El campo contiene valores alfanuméricos y se requieren valores numéricos"
                            arregloEstruc(12) = "0"
                            HuboError = True
                    End If
                    Flog.writeline
                    
                    'Dias Trabajados
                    Flog.writeline "Procesando Campo 13: Dias Trabajados"
                    'PERMITE 0   08-03-07
                    If arreglo(13) >= 0 Or arreglo(13) <= 30 Then
                        Flog.writeline Espacios(Tabulador * 1) & "Dias Trabajados Obtenido "
                        Flog.writeline
                    Else
                        Flog.writeline Espacios(Tabulador * 1) & "ERROR Los dias trabajados deben ser positivos y menores a 30"
                        HuboError = True
                    End If
                    
                    'Tipo de Linea
                    'Fijo 00 (Linea principal Default)
                    Flog.writeline "Procesando Campo 14: Tipo de Linea"
                    TipoLinea = "00"
                    Flog.writeline Espacios(Tabulador * 1) & " Tipo Linea Obtenido "
                    Flog.writeline
                    
                    'Codigo de Movimiento del Personal
                    'PERMITE 0   08-03-07
                    Flog.writeline "Procesando Campo 15: Codigo de Movimiento del Personal"
                    
                    aux = 1
                    Do While (aux <= total_mov) And (Not HuboError)
                        If IsNumeric(arregloMov(aux)) = True Then
                                Flog.writeline Espacios(Tabulador * 1) & "CAMPO Obtenido "
                        Else
                                Flog.writeline Espacios(Tabulador * 1) & "ERROR. El campo contiene valores alfanuméricos y se requieren valores numéricos"
                                HuboError = True
                                arregloMov(aux) = 0
                        End If
                        Flog.writeline
                        
                        
                        If cambioContrato Then ' si hubo cambio de contrato no toco la fecha
                        Else
                        'Fecha Desde
                        'Si Movimiento de personal es 1,3,4,5,6,7,8 el campo es obligatorio. QUEDA PARA REVISAR
                            Flog.writeline "Procesando Campo 16: Fecha Desde"
                            'FGZ - 31/05/2013 ------------------------------------------------------
                            'If arregloMov(aux) = "1" Or arregloMov(aux) = "3" Or arregloMov(aux) = "4" Or arregloMov(aux) = "5" Or arregloMov(aux) = "6" Or arregloMov(aux) = "7" Or arregloMov(aux) = "8" Then
                            If arregloMov(aux) = "1" Or arregloMov(aux) = "3" Or arregloMov(aux) = "4" Or arregloMov(aux) = "5" Or arregloMov(aux) = "6" Or arregloMov(aux) = "7" Or arregloMov(aux) = "8" Or arregloMov(aux) = "11" Then
                            'FGZ - 31/05/2013 ------------------------------------------------------
                                Flog.writeline Espacios(Tabulador * 1) & "Campo Obligatorio. "
                                Flog.writeline Espacios(Tabulador * 1) & arregloFecD(aux)
                            Else
                                Flog.writeline Espacios(Tabulador * 1) & "Campo Optativo. "
                                'Cuando es un retiro, la fecha del retiro se guarda como desde en RH Pro pero en Previred se informa en Hasta
                                If arregloMov(aux) = "2" Then
                                    arregloFecH(aux) = arregloFecD(aux)
                                End If
                                arregloFecD(aux) = vbNull
                            End If
                            Flog.writeline
                        
                        'Fecha Hasta
                        'Si Movimiento de personal es 2,3,4,6 el campo es obligatorio. QUEDA PARA REVISAR
                            Flog.writeline "Procesando Campo 17: Fecha Hasta"
                            'FGZ - 31/05/2013 ------------------------------------------------------
                            'If arregloMov(aux) = "2" Or arregloMov(aux) = "3" Or arregloMov(aux) = "4" Or arregloMov(aux) = "6" Then
                            If arregloMov(aux) = "2" Or arregloMov(aux) = "3" Or arregloMov(aux) = "4" Or arregloMov(aux) = "6" Or arregloMov(aux) = "11" Then
                            'FGZ - 31/05/2013 ------------------------------------------------------
                                Flog.writeline "Procesando Campo 17: Campo Obligatorio."
                                Flog.writeline Espacios(Tabulador * 1) & arregloFecH(aux)
                            Else
                                Flog.writeline Espacios(Tabulador * 1) & " Campo Optativo. "
                                arregloFecH(aux) = vbNull
                            End If
                            Flog.writeline
                            'aux = aux + 1
                        End If
                        aux = aux + 1
                    Loop
                    
                    '---------------------------tramo asignacion familiar--------------------------------------------------------
                    'Tramo de Asignacion Familiar
                    'Flog.writeline "Procesando Campo 18: Tramo de Asignacion Familiar"
                    'If arreglo(18) = 0 And arregloEstruc(18) = "0" Then
                    '    Flog.writeline Espacios(Tabulador * 1) & "No se encontro el codigo para el  Tramo de Asignacion Familiar"
                    '    arregloEstruc(18) = "D"
                        'HuboError = True
                    'Else
                    '    If arreglo(18) <> 0 Then
                    '        Select Case arreglo(18)
                    '            Case 1:
                    '                arregloEstruc(18) = "A"
                    '            Case 2:
                    '                arregloEstruc(18) = "B"
                    '            Case 3:
                    '                arregloEstruc(18) = "C"
                    '            Case 4:
                    '                arregloEstruc(18) = "D"
                    '            Case Else
                    '                arregloEstruc(18) = "Sin"
                    '        End Select
                    '    End If
                    '    Flog.writeline Espacios(Tabulador * 1) & "Campo Obtenido"
                    'End If
                    'Flog.writeline
                    '----------------------------hasta aca era tramo asignacion familiar -----------------------------------------
                    
                    
                    'nuevo tramo de asignacion familiar--------------------------------------------------
                     
                    Flog.writeline "Procesando Campo 18: Tramo de Asignacion Familiar"
'                    If arregloEstruc(18) = "-1" Then
'                        Flog.writeline Espacios(Tabulador * 1) & "No se encontro el codigo para el  Tramo de Asignacion Familiar"
'                        arregloEstruc(18) = "D"
'                       HuboError = True
'                    Else
                        If arreglo(18) = 0 Then
                            arregloEstruc(18) = "D"
                            Flog.writeline Espacios(Tabulador * 1) & "No se encontro el codigo para el  Tramo de Asignacion Familiar o el valor fue cero"
                        Else
                            If arreglo(18) = 1 Then
                                arregloEstruc(18) = "A"
                            Else
                                If arreglo(18) = 2 Then
                                    arregloEstruc(18) = "B"
                                Else
                                    If arreglo(18) >= 3 Then
                                        arregloEstruc(18) = "C"
                                    End If
                                End If
                            End If
                       End If
 '                  End If
                    'hasta aca nuevo tramo de asig familiar----------------------------------------------
                    
                    'Num Cargas Simples
                    Flog.writeline "Procesando Campo 19: Numero de Cargas Simples"
                    
                    If arreglo(19) = 0 Then
                        Flog.writeline Espacios(Tabulador * 1) & "Campo Vacio. OPTATIVO..CONTINUA PROCESAMIENTO"
                    Else
                        If arreglo(19) < 0 Or arreglo(19) > 9 Then
                            Flog.writeline Espacios(Tabulador * 1) & "ERROR El Numero de Cargas Simples debe ser positivos y menores a 9"
                            HuboError = True
                        Else
                            Flog.writeline Espacios(Tabulador * 1) & "Numero de Cargas Simples Obtenido "
                        End If
                    End If
                    Flog.writeline
                    
                    'Num Cargas Maternales
                    Flog.writeline "Procesando Campo 20: Numero de Cargas Maternales"
                    If arreglo(20) = 0 Then
                        Flog.writeline Espacios(Tabulador * 1) & "Campo Vacio. OPTATIVO..CONTINUA PROCESAMIENTO"
                    Else
                        If arreglo(20) < 0 Or arreglo(20) > 9 Then
                            Flog.writeline Espacios(Tabulador * 1) & "ERROR El Numero de Cargas Maternales debe ser positivos y menores a 9"
                            HuboError = True
                        Else
                            Flog.writeline Espacios(Tabulador * 1) & "Numero de Cargas Maternales Obtenido"
                        End If
                    End If
                    Flog.writeline
                    
                    'Num Cargas Invalidas
                    Flog.writeline "Procesando Campo 21: Numero de Cargas Invalidas"
                    If arreglo(21) = 0 Then
                        Flog.writeline Espacios(Tabulador * 1) & " Campo Vacio. OPTATIVO..CONTINUA PROCESAMIENTO"
                    Else
                        If arreglo(21) < 0 Or arreglo(21) > 9 Then
                            Flog.writeline Espacios(Tabulador * 1) & "ERROR El Numero de Cargas Invalidas debe ser positivos y menores a 9"
                            HuboError = True
                        Else
                            Flog.writeline Espacios(Tabulador * 1) & "Numero de Cargas Invalidas Obtenido"
                        End If
                    End If
                    Flog.writeline
                    
                    'Asignacion Familiar
                    Flog.writeline "Procesando Campo 22: Asignacion Familiar"
                    If arreglo(22) = 0 Then
                        Flog.writeline Espacios(Tabulador * 1) & "Campo Vacio. OPTATIVO..CONTINUA PROCESAMIENTO"
                    Else
                        If arreglo(21) < 0 Or arreglo(22) > 600000 Then
                            Flog.writeline Espacios(Tabulador * 1) & "ERROR Asignacion Familiar debe ser positiva y menores a 600.000"
                            HuboError = True
                        Else
                            Flog.writeline Espacios(Tabulador * 1) & "Asignacion Familiar Obtenida"
                        End If
                    End If
                    Flog.writeline
                    
                    'Asignacion Retroactiva
                    Flog.writeline "Procesando Campo 23: Asignacion Familiar Retroactiva"
                    If arreglo(23) = 0 Then
                        Flog.writeline Espacios(Tabulador * 1) & " Campo Vacio. OPTATIVO..CONTINUA PROCESAMIENTO"
                    Else
                        If arreglo(23) < 0 Or arreglo(23) > 600000 Then
                            Flog.writeline Espacios(Tabulador * 1) & "ERROR Asignacion Retroactiva debe ser positiva y menores a 600.000"
                            HuboError = True
                        Else
                            Flog.writeline Espacios(Tabulador * 1) & "Asignacion Retroactiva Obtenida"
                        End If
                    End If
                    Flog.writeline
                    
                    'Reintegro Cargas
                    Flog.writeline "Procesando Campo 24: Reintegro Cargas Familiares"
                    If arreglo(24) = 0 Then
                        Flog.writeline Espacios(Tabulador * 1) & "Campo Vacio. OPTATIVO..CONTINUA PROCESAMIENTO"
                    Else
                        If arreglo(24) < 0 Or arreglo(24) > 600000 Then
                            Flog.writeline Espacios(Tabulador * 1) & "ERROR Reintegro Cargas debe ser positivo y menores a 600.000"
                            HuboError = True
                        Else
                            Flog.writeline Espacios(Tabulador * 1) & "Reintegro Cargas Obtenida"
                        End If
                    End If
                    Flog.writeline
                    
                    'Solicitud Subsidio Trabajador Joven
                    'Fijo "N"
                    Flog.writeline "Procesando Campo 25: Solicitud Subsidio Trabajador Joven"
                    SolicSubsidioJoven = "N"
                    Flog.writeline

                    '-----------------------------------------------------------------------------------
                    ' Bloque de AFP
                    '-----------------------------------------------------------------------------------
                    
                    'Codigo de la AFP
                    Flog.writeline "Procesando Campo 26: Codigo de la AFP"
                    If arregloEstruc(26) = "0" And arreglo(26) = 0 Then
                        Flog.writeline Espacios(Tabulador * 1) & "ERROR No se encontro el codigo para la AFP"
                        HuboError = True
                    Else
                        If (IsNumeric(arregloEstruc(26)) = True) Or (IsNumeric(arreglo(26)) = True) Then
                            Flog.writeline Espacios(Tabulador * 1) & "Codigo de la AFP Obtenido "
                        Else
                            Flog.writeline Espacios(Tabulador * 1) & "ERROR. El campo contiene valores alfanuméricos y se requieren valores numéricos"
                            arreglo(26) = 0
                            arregloEstruc(26) = "0"
                            HuboError = True
                        End If
                    End If
                    Flog.writeline
                    
                    'Renta Imponible AFP
                    Flog.writeline "Procesando Campo 27: Renta Imponible AFP"
                    If arreglo(27) = 0 Then
                        Flog.writeline Espacios(Tabulador * 1) & "Campo Vacio. OPTATIVO..CONTINUA PROCESAMIENTO"
                    Else
                        'If arreglo(27) < 0 Or arreglo(27) > 600000 Then
                        If arreglo(27) < 0 Then
                            Flog.writeline Espacios(Tabulador * 1) & "ERROR Reintegro Cargas debe ser positivo"
                            HuboError = True
                        Else
                            Flog.writeline Espacios(Tabulador * 1) & "Reintegro Cargas Obtenida"
                        End If
                    End If
                    Flog.writeline
                    
                    'Cotizacion Obligatoria AFP
                    Flog.writeline "Procesando Campo 28: Cotizacion Obligatoria AFP"
                    If arreglo(28) < 0 Then
                        Flog.writeline Espacios(Tabulador * 1) & "ERROR la cotizacion no puede ser negativa"
                        HuboError = True
                    Else
                        Flog.writeline Espacios(Tabulador * 1) & "Cotizacion Obligatoria AFP Obtenido "
                    End If
                    Flog.writeline
                    
                    'Aporte Seguro Invalidez y Sobrevivencia (SIS)
                    Flog.writeline "Procesando Campo 29: Aporte Seguro Invalidez y Sobrevivencia (SIS)"
                    If arreglo(29) < 0 Then
                        Flog.writeline Espacios(Tabulador * 1) & "ERROR la cotizacion no puede ser negativa"
                        HuboError = True
                    Else
                        Flog.writeline Espacios(Tabulador * 1) & "Cotizacion Obligatoria AFP Obtenido "
                    End If
                    Flog.writeline
                    
                    'Cuenta de Ahorro Voluntario AFP
                    Flog.writeline "Procesando Campo 30: Cuenta de Ahorro Voluntario AFP"
                    If arreglo(30) = 0 Then
                        Flog.writeline Espacios(Tabulador * 1) & "Campo Vacio. OPTATIVO..CONTINUA PROCESAMIENTO"
                    Else
                        If arreglo(30) < 0 Then
                            Flog.writeline Espacios(Tabulador * 1) & "ERROR Cuenta de Ahorro Voluntario AFP no puede ser negativa"
                            HuboError = True
                        Else
                            Flog.writeline Espacios(Tabulador * 1) & "Cuenta de Ahorro Voluntario AFP Obtenido "
                        End If
                    End If
                    Flog.writeline
                    
                    'Renta Imp. Sustitutiva AFP
                    Flog.writeline "Procesando Campo 31: Renta Imp. Sustitutiva AFP"
                    If arreglo(31) = 0 Then
                        Flog.writeline Espacios(Tabulador * 1) & " Campo Vacio. OPTATIVO..CONTINUA PROCESAMIENTO"
                    Else
                        If arreglo(31) < 0 Then
                            Flog.writeline Espacios(Tabulador * 1) & "ERROR Renta Imp. Sustitutiva AFP no puede ser negativo"
                            HuboError = True
                        Else
                            Flog.writeline Espacios(Tabulador * 1) & "Renta Imp. Sustitutiva AFP Obtenida "
                        End If
                    End If
                    Flog.writeline
                    
                    'Tasa Pactada
                    Flog.writeline "Procesando Campo 32: Tasa Pactada"
                    If arreglo(32) = 0 Then
                        Flog.writeline Espacios(Tabulador * 1) & "Campo Vacio. OPTATIVO..CONTINUA PROCESAMIENTO"
                    Else
                        If arreglo(32) < 0 Then
                            Flog.writeline Espacios(Tabulador * 1) & "ERROR Tasa Pactada no puede ser negativo"
                            HuboError = True
                        Else
                            Flog.writeline Espacios(Tabulador * 1) & "Tasa Pactada Obtenida "
                        End If
                    End If
                    Flog.writeline
                    
                    'Aporte Indemnizacion Sustitutiva
                    Flog.writeline "Procesando Campo 33: Aporte Indemnizacion Sustitutiva"
                    If arreglo(33) = 0 Then
                        Flog.writeline Espacios(Tabulador * 1) & "Campo Vacio. OPTATIVO..CONTINUA PROCESAMIENTO"
                    Else
                        If arreglo(33) < 0 Then
                            Flog.writeline Espacios(Tabulador * 1) & "ERROR Aporte Indemnizacion Sustitutiva no puede ser negativa"
                            HuboError = True
                        Else
                            Flog.writeline Espacios(Tabulador * 1) & "Aporte Indemnizacion Sustitutiva Obtenido "
                        End If
                    End If
                    Flog.writeline
                    
                    'Numero de Periodos Sustitutiva
                    Flog.writeline "Procesando Campo 34: Numero de Periodos Sustitutiva"
                    Flog.writeline Espacios(Tabulador * 1) & "Campo Optativo. NO DEFINIDO. Se pone a cero"
                    Flog.writeline
                    
                    'Periodos Desde Sustitutiva
                    Flog.writeline "Procesando Campo 35: Periodos Desde Sustitutiva"
                    Flog.writeline Espacios(Tabulador * 1) & "Campo Optativo. NO DEFINIDO. Se pone a nulo"
                    Flog.writeline
                    
                    'Periodos Hasta Sustitutiva
                    Flog.writeline "Procesando Campo 36: Periodos Hasta Sustitutiva"
                    Flog.writeline Espacios(Tabulador * 1) & "Campo Optativo. NO DEFINIDO. Se pone a nulo"
                    Flog.writeline
                    
                    'Puesto de Trabajo Pesado
                    Flog.writeline "Procesando Campo 37: Puesto de Trabajo Pesado"
                    If arregloEstruc(37) = "0" Then
                        Flog.writeline Espacios(Tabulador * 1) & "Campo Vacio. OPTATIVO..CONTINUA PROCESAMIENTO"
                        arregloEstruc(37) = ""
                    Else
                        Flog.writeline Espacios(Tabulador * 1) & "Puesto de Trabajo Pesado Obtenido "
                    End If
                    Flog.writeline
                    
                    '% Cotizacion Trabajo Pesado
                    Flog.writeline "Procesando Campo 38: % Cotizacion Trabajo Pesado"
                    If arreglo(38) < 0 Then
                        Flog.writeline Espacios(Tabulador * 1) & "ERROR el porcentaje cotizacion no puede ser negativo"
                        HuboError = True
                    Else
                        Flog.writeline Espacios(Tabulador * 1) & "% Cotizacion Trabajo Pesado Obtenida "
                    End If
                    Flog.writeline
                    
                    'Cotizacion Trabajo Pesado
                    Flog.writeline "Procesando Campo 39:  Cotizacion Trabajo Pesado"
                    If arreglo(39) < 0 Then
                        Flog.writeline Espacios(Tabulador * 1) & "ERROR la cotizacion no puede ser negativa"
                        HuboError = True
                    Else
                        Flog.writeline Espacios(Tabulador * 1) & "Cotizacion Trabajo Pesado Obtenida "
                    End If
                    Flog.writeline
                    
                    '-----------------------------------------------------------------------------------
                    'Bloque de APVI
                    '-----------------------------------------------------------------------------------
                    
                   'Codigo de la Institucion APVI
                   'Validado arriba en confrep
                   
                   'Numero de Contrato APVI
                   If total_APVI > 0 Then
                        Flog.writeline "Procesando Campo 41: Numero de Contrato APVI"
                        If (arregloEstruc(41) = "0") Or (arregloEstruc(41) = "") Then
                            Flog.writeline Espacios(Tabulador * 1) & "Campo Vacio."
                        Else
                            If IsNumeric(arregloEstruc(41)) = True Then
                                Flog.writeline Espacios(Tabulador * 1) & "Numero de Contrato APVI Obtenido"
                            Else
                                Flog.writeline Espacios(Tabulador * 1) & "ERROR. El campo contiene valores alfanuméricos y se requieren valores numéricos"
                                arregloEstruc(41) = "0"
                                HuboError = True
                            End If
                        End If
                        Flog.writeline
                   Else
                        arregloEstruc(41) = "0"
                   End If
                   
                   'Forma de PAGO APVI
                   Flog.writeline "Procesando Campo 42: Forma de pago Ahorro Previsional Voluntario Individual"
                   If total_APVI > 0 Then
                        If arregloEstruc(42) = "0" And arregloEstruc(42) <> "000" Then
                             Flog.writeline Espacios(Tabulador * 1) & "ERROR. Campo Vacio. OBLIGATORIO"
                             HuboError = True
                             total_APVI = 0
                         Else
                             If IsNumeric(arregloEstruc(42)) = True Then
                                 Flog.writeline Espacios(Tabulador * 1) & "Forma de pago Ahorro Previsional Voluntario Individual Obtenido"
                             Else
                                 Flog.writeline Espacios(Tabulador * 1) & "ERROR. El campo contiene valores alfanuméricos y se requieren valores numéricos"
                                 arregloEstruc(42) = "0"
                                 HuboError = True
                                 total_APVI = 0
                             End If
                         End If
                         Flog.writeline
                   Else
                        arregloEstruc(42) = "0"
                   End If
                   
                   
                   
                    'Cotizacion Ahorro Previsional Voluntario Individual
                    'Validado arriba en confrep
                   
                   'Cotizacion Depositos Convenidos
                   'Validado arriba en confrep
                    
                    '-----------------------------------------------------------------------------------
                    'Bloque de APVC
                    '-----------------------------------------------------------------------------------
                    
                   'Codigo de la Institucion Autorizada APVC
                   'Validado arriba en confrep
                   
                   'Numero de Contrato APVC
                   If total_APVC > 0 Then
                        Flog.writeline "Procesando Campo 46: Numero de Contrato APVC"
                        If (arregloEstruc(46) = "0") Or (arregloEstruc(46) = "") Then
                            Flog.writeline Espacios(Tabulador * 1) & "Campo Vacio."
                        Else
                            If IsNumeric(arregloEstruc(46)) = True Then
                                Flog.writeline Espacios(Tabulador * 1) & "Numero de Contrato APVC Obtenido"
                            Else
                                Flog.writeline Espacios(Tabulador * 1) & "ERROR. El campo contiene valores alfanuméricos y se requieren valores numéricos"
                                arregloEstruc(46) = "0"
                                HuboError = True
                            End If
                        End If
                        Flog.writeline
                    Else
                        arregloEstruc(46) = ""
                    End If
                   
                   'Forma de PAGO APVC
                   If total_APVC > 0 Then
                        Flog.writeline "Procesando Campo 47: Forma de pago Ahorro Previsional Voluntario Colectivo"
                        If arregloEstruc(47) = "0" And arregloEstruc(47) <> "000" Then
                             Flog.writeline Espacios(Tabulador * 1) & "ERROR. Campo Vacio. OBLIGATORIO"
                             HuboError = True
                         Else
                             If IsNumeric(arregloEstruc(47)) = True Then
                                 Flog.writeline Espacios(Tabulador * 1) & "Forma de pago Ahorro Previsional Voluntario Colectivo Obtenido"
                             Else
                                 Flog.writeline Espacios(Tabulador * 1) & "ERROR. El campo contiene valores alfanuméricos y se requieren valores numéricos"
                                 arregloEstruc(47) = "0"
                                 HuboError = True
                             End If
                         End If
                         Flog.writeline
                   Else
                        arregloEstruc(47) = "0"
                   End If
                   
                    'Cotizacion Trabajador APVC
                    'Validado arriba en confrep
                    
                    'Cotizacion Empleador APVC
                    'Validado arriba en confrep
                    
                    
                    '-----------------------------------------------------------------------------------
                    'Bloque de Afiliado Voluntario
                    '-----------------------------------------------------------------------------------
                    
                    
                    'Busco el RUT y el DV del Afiliado Voluntario
                    Flog.writeline "Procesando Campo 50:  RUT Afiliado Voluntario"
                    Flog.writeline "Procesando Campo 51:  DV Afiliado Voluntario"
                    'Apellido Paterno del Afiliado Voluntario
                    Flog.writeline "Procesando Campo 52:  Apellido Paterno Afiliado Voluntario"
                    'Apellidio Materno del Afiliado Voluntario
                    Flog.writeline "Procesando Campo 53:  Apellido Materno Afiliado Voluntario"
                    'Nombres del Afiliado Voluntario
                    Flog.writeline "Procesando Campo 54:  Nombres Afiliado Voluntario"
                    'Codigo Movimiento Personal
                    Flog.writeline "Procesando Campo 55:  Codigo Movimiento Personal"
                    'Fecha Desde
                    Flog.writeline "Procesando Campo 56:  Fecha Desde"
                    'Fecha Hasta
                    Flog.writeline "Procesando Campo 57:  Fecha Hasta"
                    'Codigo de la AFP
                    Flog.writeline "Procesando Campo 58:  Codigo de la AFP"
                    'Monto Capitalizacion Voluntaria
                    Flog.writeline "Procesando Campo 59:  Monto Capitalizacion Voluntaria"
                    'Monto Ahorro Voluntario
                    Flog.writeline "Procesando Campo 60:  Monto Ahorro Voluntario"
                    'Numero de periodos de cotizacion
                    Flog.writeline "Procesando Campo 61:  Numero de periodos de cotizacion"
                    Flog.writeline Espacios(Tabulador * 1) & "???"
                    Flog.writeline
                    
                    '-----------------------------------------------------------------------------------
                    'Bloque de IPS - Fonasa
                    '-----------------------------------------------------------------------------------
                    
                    'Codigo EX caja Regimen
                    Flog.writeline "Procesando Campo 62: Codigo EX caja Regimen"
                    If arregloEstruc(62) = "0" Then
                        Flog.writeline Espacios(Tabulador * 1) & "Campo Vacio. OPTATIVO..CONTINUA PROCESAMIENTO"
                    Else
                        If IsNumeric(arregloEstruc(62)) = True Then
                            Flog.writeline Espacios(Tabulador * 1) & "Codigo EX caja Regimen Obtenido "
                        Else
                            Flog.writeline Espacios(Tabulador * 1) & "ERROR. El campo contiene valores alfanuméricos y se requieren valores numéricos"
                            arregloEstruc(62) = "0"
                            HuboError = True
                        End If
                    End If
                    Flog.writeline
                    
                    'Tasa Cotizacion Ex caja de Prevision
                    Flog.writeline "Procesando Campo 63:  Tasa Cotizacion Ex caja de Prevision"
                    If arreglo(63) < 0 Then
                        Flog.writeline Espacios(Tabulador * 1) & "ERROR la Tasa de cotizacion de la Ex caja de Prevision no puede ser negativa"
                        HuboError = True
                    Else
                        Flog.writeline Espacios(Tabulador * 1) & "Tasa Cotizacion Ex caja Obtenida "
                    End If
                    Flog.writeline
                    
                   'Renta Imponible IPS
                    Flog.writeline "Procesando Campo 64:  Renta Imponible IPS"
                    If arreglo(64) < 0 Then
                        Flog.writeline Espacios(Tabulador * 1) & "ERROR Renta Imponible IPS no puede ser negativa"
                        HuboError = True
                    End If
                    
                   'Cotizacion Obligatoria IPS
                    Flog.writeline "Procesando Campo 65:  Cotizacion Obligatoria IPS"
                    If arreglo(65) < 0 Then
                        Flog.writeline Espacios(Tabulador * 1) & "ERROR la cotizacion IPS no puede ser negativa"
                        HuboError = True
                    End If
                    'EstaINP = (UCase(arregloEstruc(11)) = "INP")
                    EstaIPS = (UCase(arregloEstruc(11)) = "IPS")
                    
                    ' Si no es negativa pregunto por la otra condicion de Validacion
                    If (arregloEstruc(12) = "0" And arreglo(13) > 0) And EstaIPS Then
                        If arreglo(65) = 0 Then
                            Flog.writeline Espacios(Tabulador * 1) & "ERROR la cotizacion IPS no puede ser CERO ya que el trabajador es activo, dias trabajados es mayor a CERO y Regimen Previsional es IPS"
                        Else
                            Flog.writeline Espacios(Tabulador * 1) & "Cotizazion Obligatoria Obtenida"
                        End If
                    Else
                        Flog.writeline Espacios(Tabulador * 1) & "Cotizazion Obligatoria Obtenida"
                    End If
                    Flog.writeline
                    
                    'Renta Imponible Desahucio
                    Flog.writeline "Procesando Campo 66:  Renta Imponible Desahucio"
                    If arreglo(66) < 0 Then
                        Flog.writeline Espacios(Tabulador * 1) & "ERROR  Renta Imponible Desahucio no puede ser negativa"
                        HuboError = True
                    Else
                        If arreglo(66) > 60 Then
                            Flog.writeline Espacios(Tabulador * 1) & "ERROR la Renta Imponible Desahucio no puede ser mayor a 60"
                            HuboError = True
                        Else
                            If arreglo(66) = 0 Then
                                Flog.writeline Espacios(Tabulador * 1) & "Campo Vacio. OPTATIVO..CONTINUA PROCESAMIENTO"
                            Else
                                Flog.writeline Espacios(Tabulador * 1) & "Renta Imponible Desahucio Obtenida "
                            End If
                        End If
                    End If
                    Flog.writeline
                    
                    'Codigo Ex caja Regimen Regimen Desahucio
                    Flog.writeline "Procesando Campo 67:  Codigo Ex caja Regimen Regimen Desahucio"
                    If arregloEstruc(67) = "0" Then
                        Flog.writeline Espacios(Tabulador * 1) & "Campo Vacio. OPTATIVO..CONTINUA PROCESAMIENTO"
                    Else
                        If IsNumeric(arregloEstruc(67)) = True Then
                            Flog.writeline Espacios(Tabulador * 1) & "Codigo EX caja Regimen Obtenido "
                        Else
                            Flog.writeline Espacios(Tabulador * 1) & "ERROR. El campo contiene valores alfanuméricos y se requieren valores numéricos"
                            arregloEstruc(67) = "0"
                            HuboError = True
                        End If
                    End If
                    Flog.writeline
                    
                    'Tasa Cotizacion Desahucio Ex-Cajas de prevision
                    Flog.writeline "Procesando Campo 68:  Tasa Cotizacion Desahucio Ex-Cajas de prevision"
                    If arreglo(68) = 0 Then
                        Flog.writeline Espacios(Tabulador * 1) & "Campo Vacio. OPTATIVO..CONTINUA PROCESAMIENTO"
                    Else
                        Flog.writeline Espacios(Tabulador * 1) & "Tasa Cotizacion Desahucio Ex-Cajas de prevision "
                    End If
                    Flog.writeline
                    
                    'Cotizacion Desahucio
                    Flog.writeline "Procesando Campo 69:  Cotizacion Desahucio"
                    If arreglo(69) < 0 Then
                        Flog.writeline Espacios(Tabulador * 1) & "ERROR la Tasa de cotizacion no puede ser negativa"
                        HuboError = True
                    End If
                    If arreglo(69) = 0 Then
                        Flog.writeline Espacios(Tabulador * 1) & "Campo Vacio. OPTATIVO..CONTINUA PROCESAMIENTO"
                    Else
                        Flog.writeline Espacios(Tabulador * 1) & "Cotizacion Desahucio Obtenida"
                    End If
                    Flog.writeline
                    
                    'Cotizacion Fonasa
                    Flog.writeline "Procesando Campo 46: Cotizacion Fonasa"
                    If arreglo(70) < 0 Then
                        Flog.writeline Espacios(Tabulador * 1) & "ERROR la Tasa de cotizacion FONASA no puede ser negativa"
                        HuboError = True
                    End If
                    If arreglo(70) = 0 Then
                        Flog.writeline Espacios(Tabulador * 1) & "Campo Vacio. OPTATIVO..CONTINUA PROCESAMIENTO"
                    Else
                        Flog.writeline Espacios(Tabulador * 1) & "Cotizacion Fonasa Obtenida"
                    End If
                    Flog.writeline
                    
                    'Cotizacion Acc de Trabajo
                    Flog.writeline "Procesando Campo 71: Cotizacion Acc de Trabajo"
                    If arreglo(71) < 0 Then
                        Flog.writeline Espacios(Tabulador * 1) & "ERROR la Cotizacion Acc de Trabajo no puede ser negativa"
                        HuboError = True
                    End If
                    If arreglo(71) = 0 Then
                        Flog.writeline Espacios(Tabulador * 1) & "Campo Vacio. OPTATIVO..CONTINUA PROCESAMIENTO"
                    Else
                        Flog.writeline Espacios(Tabulador * 1) & "Cotizacion Acc de Trabajo Obtenida"
                    End If
                    Flog.writeline
                    
                    'Bonificacion Ley 15.386
                    Flog.writeline "Procesando Campo 72: Bonificacion Ley 15.386"
                    If arreglo(72) < 0 Then
                        Flog.writeline Espacios(Tabulador * 1) & "ERROR Bonificacion Ley 15.386 no puede ser negativa"
                        HuboError = True
                    End If
                    If arreglo(72) = 0 Then
                        Flog.writeline Espacios(Tabulador * 1) & "Campo Vacio. OPTATIVO..CONTINUA PROCESAMIENTO"
                    Else
                        Flog.writeline Espacios(Tabulador * 1) & "Cotizacion Acc de Trabajo Obtenida"
                    End If
                    Flog.writeline
                    
                    'Descuentos por Cargas Familiares de ISL
                    Flog.writeline "Procesando Campo 73: Descuentos por Cargas Familiares"
                    If arreglo(73) < 0 Then
                        Flog.writeline Espacios(Tabulador * 1) & "ERROR Descuentos por Cargas Familiares no puede ser negativo"
                        HuboError = True
                    End If
                    If arreglo(73) = 0 Then
                        Flog.writeline Espacios(Tabulador * 1) & "Campo Vacio. OPTATIVO..CONTINUA PROCESAMIENTO"
                    Else
                        Flog.writeline Espacios(Tabulador * 1) & "Descuentos por Cargas Familiares Obtenido"
                    End If
                    Flog.writeline
                    
                    'Bonos Gobierno
                    Flog.writeline "Procesando Campo 74: Bonos Gobierno"
                    If arreglo(74) = 0 Then
                        Flog.writeline Espacios(Tabulador * 1) & "Campo Vacio. OPTATIVO..CONTINUA PROCESAMIENTO"
                    Else
                        Flog.writeline Espacios(Tabulador * 1) & "Bonos Gobierno obtenido"
                    End If
                    Flog.writeline
                    
                    '-----------------------------------------------------------------------------------
                    'Bloque de ISAPRE
                    '-----------------------------------------------------------------------------------
                    'Codigo ISAPRE
                    Flog.writeline "Procesando Campo 75: Codigo ISAPRE"
                    If arregloEstruc(75) = "0" Then
                        Flog.writeline Espacios(Tabulador * 1) & "ERROR No se encontro el codigo ISAPRE"
                        HuboError = True
                    Else
                        If IsNumeric(arregloEstruc(75)) = True Then
                            Flog.writeline Espacios(Tabulador * 1) & "Codigo ISAPRE Obtenido "
                        Else
                            Flog.writeline Espacios(Tabulador * 1) & "ERROR. El campo contiene valores alfanuméricos y se requieren valores numéricos"
                            arregloEstruc(75) = "0"
                            HuboError = True
                        End If
                    End If
                    Flog.writeline
                    
                    'FUN DEL EMPLEADO
                    Flog.writeline "Procesando Campo 76: FUN  "
                    '28/10/2014
                    'StrSql = " SELECT nrodoc FROM tercero" & _
                    '         " INNER JOIN ter_doc  ON (tercero.ternro = ter_doc.ternro)" & _
                    '         " INNER JOIN tipodocu on tipodocu.tidnro = ter_doc.tidnro and tipodocu.tidsigla='Fun'" & _
                    '         " WHERE tercero.ternro= " & rs_Empleados!ternro
                    
                    'Inicio
                    StrSql = " SELECT nrodoc FROM tercero "
                    StrSql = StrSql & " INNER JOIN ter_doc ON ter_doc.ternro =tercero.ternro "
                    StrSql = StrSql & " INNER JOIN tipodocu ON ter_doc.tidnro = tipodocu.tidnro "
                    StrSql = StrSql & " INNER JOIN tipodocu_pais tdp ON tipodocu.tidnro = tdp.tidnro AND tdp.paisnro = 8"
                    StrSql = StrSql & " AND UPPER(tipodocu.tidsigla) = '" & UCase("Fun") & "'"
                    StrSql = StrSql & " WHERE Tercero.ternro = " & rs_Empleados!ternro
                    'fin
                    
                    OpenRecordset StrSql, rs_Rut
                    If Not rs_Rut.EOF Then
                        FUN = rs_Rut!nrodoc
                        Flog.writeline Espacios(Tabulador * 1) & "FUN Obtenido"
                    Else
                        Flog.writeline Espacios(Tabulador * 1) & "No se encontro Numero de FUN"
                        FUN = 0
                    End If
                    Flog.writeline
                    
                    'Renta imponible Isapre
                    Flog.writeline "Procesando Campo 77:  Renta imponible Isapre"
                    If arreglo(77) < 0 Then
                        Flog.writeline Espacios(Tabulador * 1) & "ERROR la Renta imponible Isapre no puede ser negativa"
                        HuboError = True
                    End If
                    If arreglo(77) = 0 Then
                        Flog.writeline Espacios(Tabulador * 1) & "Campo Vacio. OPTATIVO..CONTINUA PROCESAMIENTO"
                    Else
                        Flog.writeline Espacios(Tabulador * 1) & "Renta imponible Isapre Obtenida"
                    End If
                    Flog.writeline
                    
                    
                    'Moneda del Plan Pactado
                    Flog.writeline "Procesando Campo 78: Moneda del Plan Pactado"
'                    If IsNumeric(arreglo(78)) = True Then
'                        Flog.writeline Espacios(Tabulador * 1) & "Moneda del Plan Pactado Obtenido "
'                    Else
'                        Flog.writeline Espacios(Tabulador * 1) & "ERROR. El campo contiene valores alfanuméricos y se requieren valores numéricos"
'                        arregloEstruc(78) = "0"
'                        HuboError = True
'                    End If

                    
                    Flog.writeline
                    
                    'Cotizacion Pactada
                    Flog.writeline "Procesando Campo 79: Cotizacion Pactada"
                    If arreglo(79) < 0 Then
                        Flog.writeline Espacios(Tabulador * 1) & "ERROR el campo no puede ser negativo"
                        HuboError = True
                    End If
                    If arreglo(79) = 0 Then
                        Flog.writeline Espacios(Tabulador * 1) & "Campo Vacio. OPTATIVO..CONTINUA PROCESAMIENTO"
                    Else
                        Flog.writeline Espacios(Tabulador * 1) & "Cotizacion Pactada Obtenida"
                    End If
                    Flog.writeline
                    
'                    'FONASA No se trata como una ISAPRE HJI 25/06/08
'                    If EsFonasa = True Then
'                        Flog.writeline "Procesando Campo 80: Fonasa No Informa Cotizacion Obligatoria ISAPRE"
'                        arreglo(80) = 0
'                    Else
                        'Cotizacion Obligatoria ISAPRE
                        Flog.writeline "Procesando Campo 80: Cotizacion Obligatoria ISAPRE"
                        If arreglo(80) < 0 Then
                            Flog.writeline Espacios(Tabulador * 1) & "ERROR la Cotizacion Obligatoria no puede ser negativa"
                            HuboError = True
                        End If
                        If arreglo(80) = 0 Then
                            Flog.writeline Espacios(Tabulador * 1) & "Campo Vacio."
                            'HuboError = True
                        Else
                            Flog.writeline Espacios(Tabulador * 1) & "Cotizacion Obligatoria Isapre Obtenida"
                        End If
'                    End If
                    Flog.writeline
                    
                    'Cotizacion Adicional Voluntaria
                    Flog.writeline "Procesando Campo 81: Cotizacion Adicional Voluntaria"
                    If arreglo(81) < 0 Then
                        Flog.writeline Espacios(Tabulador * 1) & "ERROR la Cotizacion Adicional Voluntaria no puede ser negativa"
                        HuboError = True
                    End If
                    If arreglo(81) = 0 Then
                        Flog.writeline Espacios(Tabulador * 1) & "Campo Vacio. OPTATIVO..CONTINUA PROCESAMIENTO"
                    Else
                        Flog.writeline Espacios(Tabulador * 1) & "Cotizacion Adicional Voluntaria Obtenida"
                    End If
                    Flog.writeline
                    
                    'Monto Garntía Explicita de Salud - GES
                    Flog.writeline "Procesando Campo 82:  Monto Garantía Explicita de Salud - GES"
                    If arreglo(82) < 0 Then
                        Flog.writeline Espacios(Tabulador * 1) & "ERROR el Monto Garntía Explicita de Salud - GES no puede ser negativo"
                        HuboError = True
                    End If
                    If arreglo(82) = 0 Then
                        Flog.writeline Espacios(Tabulador * 1) & "Campo Vacio. OPTATIVO..CONTINUA PROCESAMIENTO"
                    Else
                        Flog.writeline Espacios(Tabulador * 1) & "Monto Garantía Explicita de Salud - GES Obtenido"
                    End If
                    Flog.writeline
                    
                    '-----------------------------------------------------------------------------------
                    'Bloque de CCAF
                    '-----------------------------------------------------------------------------------
                    
                    'Codigo CCAF
                    Flog.writeline "Procesando Campo 83: Codigo CCAF"
                    If arregloEstruc(83) = "0" Then
                        Flog.writeline Espacios(Tabulador * 1) & "Campo Vacio. OPTATIVO..CONTINUA PROCESAMIENTO"
                    Else
                        If IsNumeric(arregloEstruc(83)) = True Then
                            Flog.writeline Espacios(Tabulador * 1) & "Codigo CCAF Obtenido"
                        Else
                            Flog.writeline Espacios(Tabulador * 1) & "ERROR. El campo contiene valores alfanuméricos y se requieren valores numéricos"
                            arregloEstruc(83) = "0"
                            HuboError = True
                        End If
                    End If
                    Flog.writeline
                    
                    'Renta imponible CCAF
                    Flog.writeline "Procesando Campo 84: Renta imponible CCAF"
                    If arreglo(84) < 0 Then
                        Flog.writeline Espacios(Tabulador * 1) & "ERROR el campo no puede ser negativo"
                        HuboError = True
                    End If
                    If arreglo(84) = 0 Then
                        Flog.writeline Espacios(Tabulador * 1) & "Campo Vacio. OPTATIVO..CONTINUA PROCESAMIENTO"
                    Else
                        Flog.writeline Espacios(Tabulador * 1) & "Renta imponible CCAF Obtenida"
                    End If
                    Flog.writeline
                    
                    'Creditos Personales
                    Flog.writeline "Procesando Campo 85: Creditos Personales"
                    If arreglo(85) < 0 Then
                        Flog.writeline Espacios(Tabulador * 1) & "ERROR el campo no puede ser negativo"
                        HuboError = True
                    End If
                    If arreglo(85) = 0 Then
                        Flog.writeline Espacios(Tabulador * 1) & "Campo Vacio. OPTATIVO..CONTINUA PROCESAMIENTO"
                    Else
                        Flog.writeline Espacios(Tabulador * 1) & "Creditos Personales Obtenido"
                    End If
                    Flog.writeline
                    
                    'Descuento Dental CCAF
                    Flog.writeline "Procesando Campo 86: Descuento Dental CCAF"
                    If arreglo(86) < 0 Then
                        Flog.writeline Espacios(Tabulador * 1) & "ERROR el campo no puede ser negativo"
                        HuboError = True
                    End If
                    If arreglo(86) = 0 Then
                        Flog.writeline Espacios(Tabulador * 1) & "Campo Vacio. OPTATIVO..CONTINUA PROCESAMIENTO"
                    Else
                        Flog.writeline Espacios(Tabulador * 1) & "Descuento Dental CCAF Obtenido"
                    End If
                    Flog.writeline
                    
                    'Descuento Por Leasing
                    Flog.writeline "Procesando Campo 87: Descuento Por Leasing"
                    If arreglo(87) < 0 Then
                        Flog.writeline Espacios(Tabulador * 1) & "ERROR el campo no puede ser negativo"
                        HuboError = True
                    End If
                    If arreglo(87) = 0 Then
                        Flog.writeline Espacios(Tabulador * 1) & "Campo Vacio. OPTATIVO..CONTINUA PROCESAMIENTO"
                    Else
                        Flog.writeline Espacios(Tabulador * 1) & "Descuento Por Leasing Obtenido"
                    End If
                    Flog.writeline
                                    
                    'Descuento por Seguro de Vida
                    Flog.writeline "Procesando Campo 88: Descuento por Seguro de Vida"
                    If arreglo(88) < 0 Then
                        Flog.writeline Espacios(Tabulador * 1) & "ERROR el campo no puede ser negativo"
                        HuboError = True
                    End If
                    If arreglo(88) = 0 Then
                        Flog.writeline Espacios(Tabulador * 1) & "Campo Vacio. OPTATIVO..CONTINUA PROCESAMIENTO"
                    Else
                        Flog.writeline Espacios(Tabulador * 1) & "Descuento por Seguro de Vida Obtenido"
                    End If
                    Flog.writeline
                    
                    'Otros CCAF
                    Flog.writeline "Procesando Campo 89: Otros CCAF"
                    If arreglo(89) < 0 Then
                        Flog.writeline Espacios(Tabulador * 1) & "ERROR el campo no puede ser negativo"
                        HuboError = True
                    End If
                    If arreglo(89) = 0 Then
                        Flog.writeline Espacios(Tabulador * 1) & "Campo Vacio. OPTATIVO..CONTINUA PROCESAMIENTO"
                    Else
                        Flog.writeline Espacios(Tabulador * 1) & "Otros CCAF Obtenido"
                    End If
                    Flog.writeline
                    
                    'Cotizacion CCAF a no Afiliado a ISAPRE
                    Flog.writeline "Procesando Campo 90: Cotizacion CCAF a no Afiliado a ISAPRE"
                    If arreglo(90) < 0 Then
                        Flog.writeline Espacios(Tabulador * 1) & "ERROR el campo no puede ser negativo"
                        HuboError = True
                    End If
                    If arreglo(90) = 0 Then
                        Flog.writeline Espacios(Tabulador * 1) & "Campo Vacio. OPTATIVO..CONTINUA PROCESAMIENTO"
                    Else
                        Flog.writeline Espacios(Tabulador * 1) & "Cotizacion CCAF a no Afiliado a ISAPRE Obtenida"
                    End If
                    Flog.writeline
                    
                    'Descuento cargas Familiares CCAF
                    Flog.writeline "Procesando Campo 91: Descuento cargas Familiares CCAF"
                    If arreglo(91) = 0 Then
                        Flog.writeline Espacios(Tabulador * 1) & "Campo Vacio. OPTATIVO..CONTINUA PROCESAMIENTO"
                    Else
                        Flog.writeline Espacios(Tabulador * 1) & "Descuento cargas Familiares CCAF Obtenido"
                    End If
                    Flog.writeline
                    
                    'Otros CCAF 1
                    Flog.writeline "Procesando Campo 92: Otros CCAF 1"
                    If arreglo(92) < 0 Then
                        Flog.writeline Espacios(Tabulador * 1) & "ERROR el campo no puede ser negativo"
                        HuboError = True
                    End If
                    If arreglo(92) = 0 Then
                        Flog.writeline Espacios(Tabulador * 1) & "Campo Vacio. OPTATIVO..CONTINUA PROCESAMIENTO"
                    Else
                        Flog.writeline Espacios(Tabulador * 1) & "Otros CCAF 1 Obtenido"
                    End If
                    Flog.writeline
                    
                    'Otros CCAF 2
                    Flog.writeline "Procesando Campo 92: Otros CCAF 2"
                    If arreglo(92) < 0 Then
                        Flog.writeline Espacios(Tabulador * 1) & "ERROR el campo no puede ser negativo"
                        HuboError = True
                    End If
                    If arreglo(92) = 0 Then
                        Flog.writeline Espacios(Tabulador * 1) & "Campo Vacio. OPTATIVO..CONTINUA PROCESAMIENTO"
                    Else
                        Flog.writeline Espacios(Tabulador * 1) & "Otros CCAF 2 Obtenido"
                    End If
                    Flog.writeline
                    
                    'Bonos Gobierno
                    Flog.writeline "Procesando Campo 94: Bonos Gobierno"
                    If arreglo(94) = 0 Then
                        Flog.writeline Espacios(Tabulador * 1) & "Campo Vacio. OPTATIVO..CONTINUA PROCESAMIENTO"
                    Else
                        Flog.writeline Espacios(Tabulador * 1) & "Bonos Gobierno obtenido"
                    End If
                    Flog.writeline
                    
                    'Codigo de Sucursal
                    Flog.writeline "Procesando Campo 95: Codigo de Sucursal"
                    If arregloEstruc(95) = "0" Then
                        Flog.writeline Espacios(Tabulador * 1) & "Campo Vacio. OPTATIVO..CONTINUA PROCESAMIENTO"
                    Else
                        If IsNumeric(arregloEstruc(95)) = True Then
                            Flog.writeline Espacios(Tabulador * 1) & "Codigo de Sucursal Obtenido"
                        Else
                            Flog.writeline Espacios(Tabulador * 1) & "ERROR. El campo contiene valores alfanuméricos y se requieren valores numéricos"
                            arregloEstruc(95) = "0"
                            HuboError = True
                        End If
                    End If
                    Flog.writeline
                    
                    '-----------------------------------------------------------------------------------
                    'Bloque de Mutual de Seguridad
                    '-----------------------------------------------------------------------------------
                    
                    'Codigo Mutual
                    Flog.writeline "Procesando Campo 96: Codigo Mutual"
                    If arregloEstruc(96) = "0" Then
                        Flog.writeline Espacios(Tabulador * 1) & "Campo Vacio. OPTATIVO..CONTINUA PROCESAMIENTO"
                    Else
                        If IsNumeric(arregloEstruc(96)) = True Then
                            Flog.writeline Espacios(Tabulador * 1) & "Codigo Mutual Obtenido"
                        Else
                            Flog.writeline Espacios(Tabulador * 1) & "ERROR. El campo contiene valores alfanuméricos y se requieren valores numéricos"
                            arregloEstruc(96) = "0"
                            HuboError = True
                        End If
                    End If
                    Flog.writeline
                    
                    'Renta Imponible Mututal
                    Flog.writeline "Procesando Campo 97:  Renta Imponible Mututal"
                    If arreglo(97) < 0 Then
                        Flog.writeline Espacios(Tabulador * 1) & "ERROR Renta Imponible Mututal no puede ser negativa"
                        HuboError = True
                    End If
                    
                    'Cotizacion ACC de Trabajo
                    Flog.writeline "Procesando Campo 98: Cotizacion ACC de Trabajo"
                    If arreglo(98) < 0 Then
                        Flog.writeline Espacios(Tabulador * 1) & "ERROR el campo no puede ser negativo"
                        HuboError = True
                    End If
                    If arreglo(98) = 0 Then
                        Flog.writeline Espacios(Tabulador * 1) & "Campo Vacio. OPTATIVO..CONTINUA PROCESAMIENTO"
                    Else
                        Flog.writeline Espacios(Tabulador * 1) & "Cotizacion ACC de Trabajo Obtenido"
                    End If
                    Flog.writeline
                    
                   'Sucursal para pago mutual
                    Flog.writeline "Procesando Campo 99: Sucursal para pago mutual"
                    If arregloEstruc(99) = "0" Then
                        Flog.writeline Espacios(Tabulador * 1) & "Campo Vacio. OPTATIVO..CONTINUA PROCESAMIENTO"
                    Else
                        If IsNumeric(arregloEstruc(99)) = True Then
                            Flog.writeline Espacios(Tabulador * 1) & "Sucursal para pago mutual Obtenida"
                        Else
                            Flog.writeline Espacios(Tabulador * 1) & "ERROR. El campo contiene valores alfanuméricos y se requieren valores numéricos"
                            arregloEstruc(99) = "0"
                            HuboError = True
                        End If
                    End If
                    Flog.writeline
                                     
                   'Renta imponible Seguro Cesantia
                    Flog.writeline "Procesando Campo 100: Renta imponible Seguro Cesantia"
                    If arreglo(100) < 0 Then
                        Flog.writeline Espacios(Tabulador * 1) & "ERROR el campo no puede ser negativo"
                        HuboError = True
                    End If
                    SeguroCesantia = False
                    If arreglo(100) = 0 Then
                        Flog.writeline Espacios(Tabulador * 1) & "Campo Vacio. OPTATIVO..CONTINUA PROCESAMIENTO"
                    Else
                        Flog.writeline Espacios(Tabulador * 1) & "Renta imponible Seguro Cesantia Obtenido"
                        SeguroCesantia = True
                    End If
                    Flog.writeline
                    
                    '-----------------------------------------------------------------------------------
                    'Bloque de Seguro Cesantia
                    '-----------------------------------------------------------------------------------
                    
                    'Aporte Trabajador Seguro Cesantia
                    Flog.writeline "Procesando Campo 101: Aporte Trabajador Seguro Cesantia"
                    If arreglo(101) < 0 Then
                        Flog.writeline Espacios(Tabulador * 1) & "ERROR el campo no puede ser negativo"
                        HuboError = True
                    End If
                    If arreglo(101) = 0 Then
                        Flog.writeline Espacios(Tabulador * 1) & "Campo Vacio. OPTATIVO..CONTINUA PROCESAMIENTO"
                    Else
                        Flog.writeline Espacios(Tabulador * 1) & "Aporte Trabajador Seguro Cesantia Obtenido"
                    End If
                    Flog.writeline
                    
                    'Aporte Empleador Seguro Cesantia
                    Flog.writeline "Procesando Campo 102: Aporte Empleador Seguro Cesantia"
                    If arreglo(102) < 0 Then
                        Flog.writeline Espacios(Tabulador * 1) & "ERROR el campo no puede ser negativo"
                        HuboError = True
                    End If
                    If arreglo(102) = 0 Then
                        Flog.writeline Espacios(Tabulador * 1) & "Campo Vacio. OPTATIVO..CONTINUA PROCESAMIENTO"
                    Else
                        Flog.writeline Espacios(Tabulador * 1) & "Aporte Empleador Seguro Cesantia Obtenido"
                    End If
                    Flog.writeline
                    
                    '-----------------------------------------------------------------------------------
                    'Bloque de Pagador de Subsidios
                    '-----------------------------------------------------------------------------------
                    
                    'Busco el RUT y el DV de Pagadpr de Subsidios
                    Flog.writeline "Procesando Campo 103:  RUT"
                    Flog.writeline "Procesando Campo 104:  DV"
                    Flog.writeline Espacios(Tabulador * 1) & "Campos Optativos. NO DEFINIDOS.Se ponen a nulo"
                    Flog.writeline
                    
                    '-----------------------------------------------------------------------------------
                    'Bloque de Datos de la Empresa
                    '-----------------------------------------------------------------------------------
                    
                    'Centro de Costos, Sucursal, Agencia, Obra, Region
                    Flog.writeline "Procesando Campo 105: Centro de Costos, Sucursal, Agencia, Obra, Region"
                    If arregloEstruc(105) = "0" Then
                        Flog.writeline Espacios(Tabulador * 1) & "Campo Vacio. OPTATIVO..CONTINUA PROCESAMIENTO"
                    Else
                        Flog.writeline Espacios(Tabulador * 1) & "Centro de Costos, Sucursal, Agencia, Obra, Region Obtenido"
                    End If
                    Flog.writeline
                    Flog.writeline
                    Flog.writeline "Fin de Procesamientos de Campos"
                    
                  'Fin Validaciones
                  '-----------------------------------------------------------------------------------
                'Controlo errores en el empleado
                'If Not HuboError Then
                    
                    Select Case TipoPago
                        Case 1, 3: 'Remuneraciones
                           'Inserto en rep_previred
                            StrSql = "INSERT INTO rep_previred (bpronro,ternro,num_linea,Titulo,pliqnro_Desde,pliqnro_hasta,empnro,rut,DV,Apellido,Apellido2,Nombres,sexo,nacionalidad,tipo_pago,Periodo_desde,Periodo_hasta,renta_imp,reg_pre,TipTrabajador,DiasTrab,TipoDeLinea,CodmovPer,fechadesde"
                            StrSql = StrSql & ",Fechahasta , TramoAsigFam, NumCargasSim, NumCargasMat, NumCargasInv, AsigFam, AsigFamRetro, ReintCarFam, SolicSubsidioTrabJoven, CodAFP, RentaImponibleAFP, CotizObligAFP, AporteSIS, CAVoluntaAFP, RenImpSustAFP, TasaPact, AportIndem, NumPeriodos, PeriDesdeAFP"
                            StrSql = StrSql & ",PeriHastaAFP,PuesTrabPesado,PorcCotizTrabPesa,CotizTrabPesa,InstAutAPV,NumContratoAPVI,ForPagAPV,CotizAPV,CotizDepConv,CodInstAutorizadaAPVC,NumContratoAPVC,FPagoAPVC,CotizTrabajadorAPVC,CotizEmpleadorAPVC,RUTAfVolunt,DVAfVolunt,ApePatVolunt"
                            StrSql = StrSql & ",ApeMatVolunt,NombVolunt,CodMovPersVolunt,FecDesdeVolunt,FecHastaVolunt,CodAFPVolunt,MontoCapVolunt,MontoAhorroVolunt,NumPerVolunt,CodCaReg,TasaCotCajPrev,RentaImpIPS,CotizObligINP,RentImpoDesah,CodCaRegDesah,TasaCotDesah,CotizDesah"
                            StrSql = StrSql & ",CotizFonasa,CotizAccTrab,BonLeyInp,DescCargFam,BonosGobierno,CodInstSal,NumFun,RentaImpIsapre,MonPlanIsapre,CotizPact,CotizObligIsapre,CotizAdicVolun,MontoGarantiaSaludGES,CodCCAF,RentaImponibleCCAF,CredPerCCAF,DescDentCCAF"
                            StrSql = StrSql & ",DescLeasCCAF,DescVidaCCAF,OtrosDesCCAF,CotCCAFnoIsapre,DesCarFamCCAF,OtrosDesCCAF1,OtrosDesCCAF2,BonosGobiernoCCAF,CodigoSucursalCCAF,CodMut,RentaimpMut,CotizAccTrabMut,SucPagMut,RentTotImp,AporTrabSeg,AporEmpSeg,RutPag,DVPag,CentroCosto,auxdeci"
                            
                            StrSql = StrSql & ") VALUES ("
                            
                            StrSql = StrSql & NroProcesoBatch & ","
                            StrSql = StrSql & rs_Empleados!ternro & ","
                            StrSql = StrSql & Num_linea & ","
                            StrSql = StrSql & "'" & Left(Titulo, 50) & "',"
                            StrSql = StrSql & "Null,"
                            StrSql = StrSql & "Null,"
                            StrSql = StrSql & Empresa & ","
                            StrSql = StrSql & "'" & RUT & "',"
                            StrSql = StrSql & "'" & DV & "',"
                            StrSql = StrSql & "'" & Apellido & "',"
                            StrSql = StrSql & "'" & Apellido2 & "',"
                            StrSql = StrSql & "'" & NombreEmp & "',"
                            StrSql = StrSql & "'" & Sexo & "',"
                            StrSql = StrSql & "'" & Nacionalidad & "',"
                            StrSql = StrSql & TipoPago & ","
                            StrSql = StrSql & "'" & Month(Fechadesde) & Year(Fechadesde) & "',"
                            StrSql = StrSql & "'" & Month(Fechahasta) & Year(Fechahasta) & "',"
                            StrSql = StrSql & arreglo(10) & "," 'Renta imponible
                            StrSql = StrSql & "'" & Mid(arregloEstruc(11), 1, 50) & "'," 'Regimen Previsional
                            StrSql = StrSql & arregloEstruc(12) & "," 'Tipo Trabajador
                            StrSql = StrSql & arreglo(13) & "," 'Dias Trabajados
                            StrSql = StrSql & TipoLinea & "," 'Tipo de Linea
                            If cambioContrato Then
                                StrSql = StrSql & 7 & "," 'Codigo Movimiento de personal
                                
                            Else
                                StrSql = StrSql & arregloMov(1) & "," 'Codigo Movimiento de personal
                            End If
                            'If cambioContrato Then
                            '    StrSql = StrSql & IIf(arregloFecD(aux - 1) = vbNull, "Null", ConvFecha(arregloFecD(aux - 1))) & "," 'Fecha Desde para el movimiento
                            '    StrSql = StrSql & IIf(arregloFecH(aux - 1) = vbNull, "Null", ConvFecha(arregloFecH(aux - 1))) & "," 'Fecha Hasta para el movimiento
                            'Else
                            StrSql = StrSql & IIf(arregloFecD(1) = vbNull, "Null", ConvFecha(arregloFecD(1))) & "," 'Fecha Desde para el movimiento
                            StrSql = StrSql & IIf(arregloFecH(1) = vbNull, "Null", ConvFecha(arregloFecH(1))) & "," 'Fecha Hasta para el movimiento
                            'End If
                            StrSql = StrSql & "'" & arregloEstruc(18) & "'," 'Tramo Asignacion Familiar
                            StrSql = StrSql & arreglo(19) & "," 'Numero de cargas simples
                            StrSql = StrSql & arreglo(20) & "," 'Numero de cargas Maternales
                            StrSql = StrSql & arreglo(21) & "," 'Numero de cargas Invalidas
                            StrSql = StrSql & arreglo(22) & "," 'ASignacion Familiar
                            StrSql = StrSql & arreglo(23) & "," 'ASignacion Familiar Retroactiva
                            StrSql = StrSql & arreglo(24) & "," 'Renta Carga Familiares
                            StrSql = StrSql & "'" & SolicSubsidioJoven & "'," 'Solicitud Subsidio Trabajador Joven
                            If arreglo(26) = 0 Then
                                StrSql = StrSql & arregloEstruc(26) & "," 'Codigo AFP
                            Else
                                StrSql = StrSql & arreglo(26) & "," 'Codigo AFP
                            End If
                            StrSql = StrSql & arreglo(27) & "," 'Renta Imponible AFP
                            StrSql = StrSql & arreglo(28) & "," 'Cotizacion Obligatoria AFP
                            StrSql = StrSql & arreglo(29) & "," 'Aporte Seguro Invalidez y Sobervivencia
                            StrSql = StrSql & arreglo(30) & "," 'Cuenta de Ahorro Voluntaria
                            StrSql = StrSql & arreglo(31) & "," 'Renta Imponible Sust a AFP
                            StrSql = StrSql & arreglo(32) & "," 'Tasa Pactada
                            StrSql = StrSql & arreglo(33) & "," 'Aporte Indem
                            StrSql = StrSql & 0 & "," 'Numero de Periodos REVISAR!!!!!!!!!!!!!!!!!!!!!
                            StrSql = StrSql & "Null" & ","  'REVISAR CAMBIAR
                            StrSql = StrSql & "Null" & ","  'REVISAR CAMBIAR
                            StrSql = StrSql & "'" & Mid(arregloEstruc(37), 1, 40) & "'," 'Puesto de trabajo Pesado
                            StrSql = StrSql & arreglo(38) & "," 'Porcentaje Cotizacion Trabajo Pesado
                            StrSql = StrSql & arreglo(39) & "," 'Cotizacion Trabajo Pesado
                            If total_APVI > 0 Then
                                StrSql = StrSql & arregloAPVI(1).Cod & "," 'Inst Autor APVI
                                StrSql = StrSql & arregloEstruc(41) & "," 'Numero de contrato APVI
                                StrSql = StrSql & arregloEstruc(42) & "," 'Forma de Pago APVI
                                StrSql = StrSql & arregloAPVI(1).Cotiza & "," 'Cotizacion APVI
                                StrSql = StrSql & arregloAPVI(1).Depositos & "," 'Cotizacion Depositos convenidos
                            Else
                                StrSql = StrSql & "0," 'Inst Autor APVI
                                StrSql = StrSql & "0," 'Numero de contrato APVI
                                StrSql = StrSql & "0," 'Forma de Pago APVI
                                StrSql = StrSql & "0," 'Cotizacion APVI
                                StrSql = StrSql & "0," 'Cotizacion Depositos convenidos
                            End If
                            If total_APVC > 0 Then
                                StrSql = StrSql & arregloAPVC(1).Cod & ","  'Codigo Institucion Autorizada APVC
                                StrSql = StrSql & arregloEstruc(46) & ","  'Numero de Contrato APVC
                                StrSql = StrSql & arregloEstruc(47) & ","  'Forma de Pago APVC
                                StrSql = StrSql & arregloAPVC(1).Cotiza & "," 'Cotizacion Trabajador APVC
                                StrSql = StrSql & arregloAPVC(1).Depositos & "," 'Cotizacion Empleador APVC
                            Else
                                StrSql = StrSql & "0,"  'Codigo Institucion Autorizada APVC
                                StrSql = StrSql & "0,"  'Numero de Contrato APVC
                                StrSql = StrSql & "0,"  'Forma de Pago APVC
                                StrSql = StrSql & "0," 'Cotizacion Trabajador APVC
                                StrSql = StrSql & "0," 'Cotizacion Empleador APVC
                            End If
                            StrSql = StrSql & "'" & "" & "'," 'RUT Afiliado Voluntario
                            StrSql = StrSql & "'" & "" & "'," 'DV Afiliado Voluntario
                            StrSql = StrSql & "'" & "" & "'," 'Apellido Paterno Afiliado Voluntario
                            StrSql = StrSql & "'" & "" & "'," 'Apellido Materno Afiliado Voluntario
                            StrSql = StrSql & "'" & "" & "'," 'Nombres Afiliado Voluntario
                            StrSql = StrSql & 0 & "," 'Codigo Movimiento Personal Afiliado Voluntario
                            StrSql = StrSql & "Null" & "," 'Fecha Desde Afiliado Voluntario
                            StrSql = StrSql & "Null" & "," 'Fecha Hasta Afiliado Voluntario
                            StrSql = StrSql & 0 & "," 'Codigo de la AFP
                            StrSql = StrSql & 0 & "," 'Monto Capitalizacion Voluntaria
                            StrSql = StrSql & 0 & "," 'Monto Ahorro Voluntario
                            StrSql = StrSql & 0 & "," 'Nu8mero de Periodos de Cotizacion
                            StrSql = StrSql & arregloEstruc(62) & "," 'Codigo Caja Regimen
                            StrSql = StrSql & arreglo(63) & "," 'Tasa cotizacion EX cajas de Regimen
                            StrSql = StrSql & arreglo(64) & "," 'Renta Imponible IPS
                            StrSql = StrSql & arreglo(65) & "," 'Cotizacion Obligatoria INP
                            StrSql = StrSql & arreglo(66) & "," 'Renta Imponible Desahucio
                            StrSql = StrSql & arregloEstruc(67) & "," 'Codigo ex caja Regimen
                            StrSql = StrSql & arreglo(68) & "," 'Tasa Cotizacion Desahucio
                            StrSql = StrSql & arreglo(69) & "," 'Cotizacion Desahucio
                            StrSql = StrSql & arreglo(70) & "," 'Cotizacion Fonasa
                            StrSql = StrSql & arreglo(71) & "," 'Cotizacion Accidente de Trabajo
                            StrSql = StrSql & arreglo(72) & "," 'Bonificacion ley 15.386
                            StrSql = StrSql & arreglo(73) & "," 'Descuento por Cargas Familiares
                            StrSql = StrSql & arreglo(74) & "," 'Bonos Gobierno
                            StrSql = StrSql & arregloEstruc(75) & "," 'Codigo Institucion de Salud
                            StrSql = StrSql & FUN & ","  'FUN
                            StrSql = StrSql & arreglo(77) & "," 'Renta Imponible Isapre
                            StrSql = StrSql & arreglo(78) & "," 'Moneda del plan pactado con Isapre
                            StrSql = StrSql & arreglo(79) & "," 'Cotizacion Pactada
                            StrSql = StrSql & arreglo(80) & "," 'Cotizacion Obligatoria ISAPRE
                            StrSql = StrSql & arreglo(81) & "," 'Cotizacion adicional Voluntaria
                            StrSql = StrSql & arreglo(82) & "," 'Monto Garantia Explicito
                            StrSql = StrSql & arregloEstruc(83) & "," 'Codigo CCAF
                            StrSql = StrSql & arreglo(84) & "," 'Renta imponible CCAF
                            StrSql = StrSql & arreglo(85) & "," 'Creditos Personales CCAF
                            StrSql = StrSql & arreglo(86) & "," 'Descuento Dental
                            StrSql = StrSql & arreglo(87) & "," 'Descuento por Leasing
                            StrSql = StrSql & arreglo(88) & "," 'Descuentos por Seguro de Vida
                            StrSql = StrSql & arreglo(89) & "," 'Otros Descuentos CCAF
                            StrSql = StrSql & arreglo(90) & "," 'Cotiz CCAF de no Afiliados ISAPRE
                            StrSql = StrSql & arreglo(91) & "," 'Descuentos Cargas Familiares CCAF
                            StrSql = StrSql & arreglo(92) & "," 'Otros Descuentos CCAF 1
                            StrSql = StrSql & arreglo(93) & "," 'Otros Descuentos CCAF 2
                            StrSql = StrSql & arreglo(94) & "," 'Bonos Gobierno
                            StrSql = StrSql & arregloEstruc(95) & "," 'Codigo de Sucursal
                            StrSql = StrSql & arregloEstruc(96) & "," 'Codigo Mutual
                            StrSql = StrSql & arreglo(97) & "," 'Renta Imponible Mutual
                            StrSql = StrSql & arreglo(98) & "," 'Cotizacion Accidente del trabajo
                            StrSql = StrSql & arregloEstruc(99) & "," 'Sucursal Para Pago Mutual
                            StrSql = StrSql & IIf(EstaIPS And SeguroCesantia, 0, arreglo(100)) & "," 'Renta total Imponible Seguro de Cesantia
                            StrSql = StrSql & IIf(EstaIPS And SeguroCesantia, 0, arreglo(101)) & "," 'Aporte Trabajador Seguro de Cesantia
                            StrSql = StrSql & IIf(EstaIPS And SeguroCesantia, 0, arreglo(102)) & "," 'Aporte Empleador Seguro de Cesantia
                            StrSql = StrSql & "0" & "," 'Rut Pagadora Subsidio
                            StrSql = StrSql & "Null" & "," 'DV Pagadora Subsidio
                            StrSql = StrSql & "'" & arregloEstruc(105) & "',"   'Centro Costo, sucursal, etc
                            StrSql = StrSql & IIf(HuboError, -1, 0)
                            StrSql = StrSql & ")"
                            Flog.writeline
                            Flog.writeline "Insertando 1 : " & StrSql
                            objConn.Execute StrSql, , adExecuteNoRecords
                            Flog.writeline
                            Flog.writeline
                            'Sumo el numero de linea
                            Num_linea = Num_linea + 1
                            CantEmplSinError = CantEmplSinError + 1
                            Flog.writeline
                            Flog.writeline "SE GRABO EL EMPLEADO "
                            Flog.writeline
                            
                            'Aperturas de lineas de movimiento
                            aux = 2 'antes 2
                            Do While aux <= total_mov And (aux < 30)
                               'Inserto en rep_previred las lineas adicionales
                                StrSql = "INSERT INTO rep_previred (bpronro,ternro,num_linea,Titulo,pliqnro_Desde,pliqnro_hasta,empnro,rut,DV,Apellido,Apellido2,Nombres,sexo,nacionalidad,tipo_pago,Periodo_desde,Periodo_hasta,renta_imp,reg_pre,TipTrabajador,DiasTrab,TipoDeLinea,CodmovPer,fechadesde"
                                StrSql = StrSql & ",Fechahasta , TramoAsigFam, NumCargasSim, NumCargasMat, NumCargasInv, AsigFam, AsigFamRetro, ReintCarFam, SolicSubsidioTrabJoven, CodAFP, RentaImponibleAFP, CotizObligAFP, AporteSIS, CAVoluntaAFP, RenImpSustAFP, TasaPact, AportIndem, NumPeriodos, PeriDesdeAFP"
                                StrSql = StrSql & ",PeriHastaAFP,PuesTrabPesado,PorcCotizTrabPesa,CotizTrabPesa,InstAutAPV,NumContratoAPVI,ForPagAPV,CotizAPV,CotizDepConv,CodInstAutorizadaAPVC,NumContratoAPVC,FPagoAPVC,CotizTrabajadorAPVC,CotizEmpleadorAPVC,RUTAfVolunt,DVAfVolunt,ApePatVolunt"
                                StrSql = StrSql & ",ApeMatVolunt,NombVolunt,CodMovPersVolunt,FecDesdeVolunt,FecHastaVolunt,CodAFPVolunt,MontoCapVolunt,MontoAhorroVolunt,NumPerVolunt,CodCaReg,TasaCotCajPrev,RentaImpIPS,CotizObligINP,RentImpoDesah,CodCaRegDesah,TasaCotDesah,CotizDesah"
                                StrSql = StrSql & ",CotizFonasa,CotizAccTrab,BonLeyInp,DescCargFam,BonosGobierno,CodInstSal,NumFun,RentaImpIsapre,MonPlanIsapre,CotizPact,CotizObligIsapre,CotizAdicVolun,MontoGarantiaSaludGES,CodCCAF,RentaImponibleCCAF,CredPerCCAF,DescDentCCAF"
                                StrSql = StrSql & ",DescLeasCCAF,DescVidaCCAF,OtrosDesCCAF,CotCCAFnoIsapre,DesCarFamCCAF,OtrosDesCCAF1,OtrosDesCCAF2,BonosGobiernoCCAF,CodigoSucursalCCAF,CodMut,RentaimpMut,CotizAccTrabMut,SucPagMut,RentTotImp,AporTrabSeg,AporEmpSeg,RutPag,DVPag,CentroCosto,auxdeci"
                                
                                StrSql = StrSql & ") VALUES ("
                                
                                StrSql = StrSql & NroProcesoBatch & ","
                                StrSql = StrSql & rs_Empleados!ternro & ","
                                StrSql = StrSql & Num_linea & ","
                                StrSql = StrSql & "'" & Left(Titulo, 50) & "',"
                                StrSql = StrSql & "Null,"
                                StrSql = StrSql & "Null,"
                                StrSql = StrSql & Empresa & ","
                                StrSql = StrSql & "'" & RUT & "',"
                                StrSql = StrSql & "'" & DV & "',"
                                StrSql = StrSql & "'" & Apellido & "',"
                                StrSql = StrSql & "'" & Apellido2 & "',"
                                StrSql = StrSql & "'" & NombreEmp & "',"
                                StrSql = StrSql & "'" & Sexo & "',"
                                StrSql = StrSql & "'" & Nacionalidad & "',"
                                StrSql = StrSql & TipoPago & ","
                                StrSql = StrSql & "'" & Month(Fechadesde) & Year(Fechadesde) & "',"
                                StrSql = StrSql & "'" & Month(Fechahasta) & Year(Fechahasta) & "',"
                                StrSql = StrSql & "0," 'Renta imponible
                                StrSql = StrSql & "'" & Mid(arregloEstruc(11), 1, 50) & "'," 'Regimen Previsional
                                StrSql = StrSql & arregloEstruc(12) & "," 'Tipo Trabajador
                                StrSql = StrSql & "0," 'Dias Trabajados
                                StrSql = StrSql & "01" & "," 'Tipo de Linea Adicional
                                StrSql = StrSql & arregloMov(aux) & "," 'Codigo Movimiento de personal
                                If cambioContrato Then
                                StrSql = StrSql & IIf(arregloFecD(aux) = vbNull, "Null", ConvFecha(arregloFecD(aux))) & "," 'Fecha Desde para el movimiento
                                StrSql = StrSql & "Null," 'Fecha Hasta para el movimiento
                                Else
                                StrSql = StrSql & IIf(arregloFecD(aux) = vbNull, "Null", ConvFecha(arregloFecD(aux))) & "," 'Fecha Desde para el movimiento
                                StrSql = StrSql & IIf(arregloFecH(aux) = vbNull, "Null", ConvFecha(arregloFecH(aux))) & "," 'Fecha Hasta para el movimiento
                                End If
                                StrSql = StrSql & "' '," 'Tramo Asignacion Familiar
                                StrSql = StrSql & "0," 'Numero de cargas simples
                                StrSql = StrSql & "0," 'Numero de cargas Maternales
                                StrSql = StrSql & "0," 'Numero de cargas Invalidas
                                StrSql = StrSql & "0," 'ASignacion Familiar
                                StrSql = StrSql & "0," 'ASignacion Familiar Retroactiva
                                StrSql = StrSql & "0," 'Renta Carga Familiares
                                StrSql = StrSql & "''," 'Solicitud Subsidio Trabajador Joven
                                If arreglo(26) = 0 Then
                                    StrSql = StrSql & arregloEstruc(26) & "," 'Codigo AFP
                                Else
                                    StrSql = StrSql & arreglo(26) & "," 'Codigo AFP
                                End If
                                'StrSql = StrSql & arregloEstruc(26) & "," 'Codigo AFP
                                StrSql = StrSql & "0," 'Renta Imponible AFP
                                StrSql = StrSql & "0," 'Cotizacion Obligatoria AFP
                                StrSql = StrSql & "0," 'Aporte Seguro Invalidez y Sobervivencia
                                StrSql = StrSql & "0," 'Cuenta de Ahorro Voluntaria
                                StrSql = StrSql & "0," 'Renta Imponible Sust a AFP
                                StrSql = StrSql & "0.00," 'Tasa Pactada
                                StrSql = StrSql & "0," 'Aporte Indem
                                StrSql = StrSql & "0," 'Numero de Periodos REVISAR!!!!!!!!!!!!!!!!!!!!!
                                StrSql = StrSql & "Null" & ","  'REVISAR CAMBIAR
                                StrSql = StrSql & "Null" & ","  'REVISAR CAMBIAR
                                StrSql = StrSql & "Null," 'Puesto de trabajo Pesado
                                StrSql = StrSql & "00.00," 'Porcentaje Cotizacion Trabajo Pesado
                                StrSql = StrSql & "0," 'Cotizacion Trabajo Pesado
                                StrSql = StrSql & "000," 'Inst Autor APVI
                                StrSql = StrSql & "0," 'Numero de contrato APVI
                                StrSql = StrSql & "0,"  'Forma de Pago APVI
                                StrSql = StrSql & "0," 'Cotizacion APVI
                                StrSql = StrSql & "0," 'Cotizacion Depositos convenidos
                                StrSql = StrSql & "0," 'Codigo Institucion Autorizada APVC
                                StrSql = StrSql & "'0'," 'Numero de Contrato APVC
                                StrSql = StrSql & "0," 'Forma de Pago APVC
                                StrSql = StrSql & "0," 'Cotizacion Trabajador APVC
                                StrSql = StrSql & "0," 'Cotizacion Empleador APVC
                                StrSql = StrSql & "'" & "" & "'," 'RUT Afiliado Voluntario
                                StrSql = StrSql & "'" & "" & "'," 'DV Afiliado Voluntario
                                StrSql = StrSql & "'" & "" & "'," 'Apellido Paterno Afiliado Voluntario
                                StrSql = StrSql & "'" & "" & "'," 'Apellido Materno Afiliado Voluntario
                                StrSql = StrSql & "'" & "" & "'," 'Nombres Afiliado Voluntario
                                StrSql = StrSql & "0," 'Codigo Movimiento Personal Afiliado Voluntario
                                StrSql = StrSql & "Null" & "," 'Fecha Desde Afiliado Voluntario
                                StrSql = StrSql & "Null" & "," 'Fecha Hasta Afiliado Voluntario
                                StrSql = StrSql & "0," 'Codigo de la AFP
                                StrSql = StrSql & "0," 'Monto Capitalizacion Voluntaria
                                StrSql = StrSql & "0," 'Monto Ahorro Voluntario
                                StrSql = StrSql & "0," 'Nu8mero de Periodos de Cotizacion
                                StrSql = StrSql & arregloEstruc(62) & "," 'Codigo Caja Regimen
                                StrSql = StrSql & "00.00," 'Tasa cotizacion EX cajas de Regimen
                                StrSql = StrSql & "0," 'Renta Imponible IPS
                                StrSql = StrSql & "0," 'Cotizacion Obligatoria INP
                                StrSql = StrSql & "0," 'Renta Imponible Desahucio
                                StrSql = StrSql & "0," 'Codigo ex caja Regimen
                                StrSql = StrSql & "0," 'Tasa Cotizacion Desahucio
                                StrSql = StrSql & "0," 'Cotizacion Desahucio
                                StrSql = StrSql & "0," 'Cotizacion Fonasa
                                StrSql = StrSql & "0," 'Cotizacion Accidente de Trabajo
                                StrSql = StrSql & "0," 'Bonificacion ley 15.386
                                StrSql = StrSql & "0," 'Descuento por Cargas Familiares
                                StrSql = StrSql & "0," 'Bonos Gobierno
                                StrSql = StrSql & arregloEstruc(75) & "," 'Codigo Institucion de Salud
                                StrSql = StrSql & "0,"  'FUN
                                StrSql = StrSql & "0," 'Renta Imponible Isapre
                                StrSql = StrSql & "0," 'Moneda del plan pactado con Isapre
                                StrSql = StrSql & "0," 'Cotizacion Pactada
                                StrSql = StrSql & "0," 'Cotizacion Obligatoria ISAPRE
                                StrSql = StrSql & "0," 'Cotizacion adicional Voluntaria
                                StrSql = StrSql & "0," 'Monto Garantia Explicito de Salud - GES
                                StrSql = StrSql & arregloEstruc(83) & "," 'Codigo CCAF
                                StrSql = StrSql & "0," 'Renta imponible CCAF
                                StrSql = StrSql & "0," 'Creditos Personales CCAF
                                StrSql = StrSql & "0," 'Descuento Dental
                                StrSql = StrSql & "0," 'Descuento por Leasing
                                StrSql = StrSql & "0," 'Descuentos por Seguro de Vida
                                StrSql = StrSql & "0," 'Otros Descuentos CCAF
                                StrSql = StrSql & "0," 'Cotiz CCAF de no Afiliados ISAPRE
                                StrSql = StrSql & "0," 'Descuentos Cargas Familiares CCAF
                                StrSql = StrSql & "0," 'Otros Descuentos CCAF 1
                                StrSql = StrSql & "0," 'Otros Descuentos CCAF 2
                                StrSql = StrSql & "0," 'Bonos Gobierno
                                StrSql = StrSql & arregloEstruc(95) & "," 'Codigo de Sucursal
                                StrSql = StrSql & arregloEstruc(96) & "," 'Codigo Mutual
                                StrSql = StrSql & "0," 'Renta Imponible Mutual
                                'StrSql = StrSql & arreglo(97) & "," 'Renta Imponible Mutual
                                StrSql = StrSql & "0," 'Cotizacion Accidente del trabajo
                                StrSql = StrSql & arregloEstruc(99) & "," 'Sucursal Para Pago Mutual
                                
                                'agrego sebastian stremel 27/01/2012
                                If cambioContrato Then
                                    StrSql = StrSql & arreglo(200) & "," 'Renta total Imponible Seguro de Cesantia
                                    StrSql = StrSql & arreglo(201) & "," 'Aporte Trabajador Seguro de Cesantia
                                    StrSql = StrSql & arreglo(202) & "," 'Aporte Empleador Seguro de Cesantia
                                Else
                                    StrSql = StrSql & "0," 'Renta total Imponible Seguro de Cesantia
                                    StrSql = StrSql & "0," 'Aporte Trabajador Seguro de Cesantia
                                    StrSql = StrSql & "0," 'Aporte Empleador Seguro de Cesantia
                                End If
                                'hasta aca
                                
                                'saco sebastian stremel 27/01/2012
                                'StrSql = StrSql & "0," 'Renta total Imponible Seguro de Cesantia
                                'StrSql = StrSql & "0," 'Aporte Trabajador Seguro de Cesantia
                                'StrSql = StrSql & "0," 'Aporte Empleador Seguro de Cesantia
                                'hasta aca
                                
                                StrSql = StrSql & "Null" & "," 'Rut Pagadora Subsidio
                                StrSql = StrSql & "Null" & "," 'DV Pagadora Subsidio
                                StrSql = StrSql & "'" & arregloEstruc(105) & "',"   'Centro Costo, sucursal, etc
                                StrSql = StrSql & IIf(HuboError, -1, 0)
                                StrSql = StrSql & ")"
                                Flog.writeline
                                Flog.writeline "Insertando 2 : " & StrSql
                                objConn.Execute StrSql, , adExecuteNoRecords
                                Flog.writeline
                                Flog.writeline
                                'Sumo el numero de linea
                                Num_linea = Num_linea + 1
                                Flog.writeline
                                Flog.writeline "SE GRABO LINEA ADICIONAL POR MOVIMIENTO"
                                Flog.writeline
            
                                aux = aux + 1
                            Loop
                            
                            'Aperturas de lineas de APVI
                            aux = 2
                            Do While (aux <= total_APVI) And (aux < 30)
                               'Inserto en rep_previred las lineas adicionales APVI
                                StrSql = "INSERT INTO rep_previred (bpronro,ternro,num_linea,Titulo,pliqnro_Desde,pliqnro_hasta,empnro,rut,DV,Apellido,Apellido2,Nombres,sexo,nacionalidad,tipo_pago,Periodo_desde,Periodo_hasta,renta_imp,reg_pre,TipTrabajador,DiasTrab,TipoDeLinea,CodmovPer,fechadesde"
                                StrSql = StrSql & ",Fechahasta , TramoAsigFam, NumCargasSim, NumCargasMat, NumCargasInv, AsigFam, AsigFamRetro, ReintCarFam, SolicSubsidioTrabJoven, CodAFP, RentaImponibleAFP, CotizObligAFP, AporteSIS, CAVoluntaAFP, RenImpSustAFP, TasaPact, AportIndem, NumPeriodos, PeriDesdeAFP"
                                StrSql = StrSql & ",PeriHastaAFP,PuesTrabPesado,PorcCotizTrabPesa,CotizTrabPesa,InstAutAPV,NumContratoAPVI,ForPagAPV,CotizAPV,CotizDepConv,CodInstAutorizadaAPVC,NumContratoAPVC,FPagoAPVC,CotizTrabajadorAPVC,CotizEmpleadorAPVC,RUTAfVolunt,DVAfVolunt,ApePatVolunt"
                                StrSql = StrSql & ",ApeMatVolunt,NombVolunt,CodMovPersVolunt,FecDesdeVolunt,FecHastaVolunt,CodAFPVolunt,MontoCapVolunt,MontoAhorroVolunt,NumPerVolunt,CodCaReg,TasaCotCajPrev,RentaImpIPS,CotizObligINP,RentImpoDesah,CodCaRegDesah,TasaCotDesah,CotizDesah"
                                StrSql = StrSql & ",CotizFonasa,CotizAccTrab,BonLeyInp,DescCargFam,BonosGobierno,CodInstSal,NumFun,RentaImpIsapre,MonPlanIsapre,CotizPact,CotizObligIsapre,CotizAdicVolun,MontoGarantiaSaludGES,CodCCAF,RentaImponibleCCAF,CredPerCCAF,DescDentCCAF"
                                StrSql = StrSql & ",DescLeasCCAF,DescVidaCCAF,OtrosDesCCAF,CotCCAFnoIsapre,DesCarFamCCAF,OtrosDesCCAF1,OtrosDesCCAF2,BonosGobiernoCCAF,CodigoSucursalCCAF,CodMut,RentaimpMut,CotizAccTrabMut,SucPagMut,RentTotImp,AporTrabSeg,AporEmpSeg,RutPag,DVPag,CentroCosto,auxdeci"
                                
                                StrSql = StrSql & ") VALUES ("
                                
                                StrSql = StrSql & NroProcesoBatch & ","
                                StrSql = StrSql & rs_Empleados!ternro & ","
                                StrSql = StrSql & Num_linea & ","
                                StrSql = StrSql & "'" & Left(Titulo, 50) & "',"
                                StrSql = StrSql & "Null,"
                                StrSql = StrSql & "Null,"
                                StrSql = StrSql & Empresa & ","
                                StrSql = StrSql & "'" & RUT & "',"
                                StrSql = StrSql & "'" & DV & "',"
                                StrSql = StrSql & "'" & Apellido & "',"
                                StrSql = StrSql & "'" & Apellido2 & "',"
                                StrSql = StrSql & "'" & NombreEmp & "',"
                                StrSql = StrSql & "'" & Sexo & "',"
                                StrSql = StrSql & "'" & Nacionalidad & "',"
                                StrSql = StrSql & TipoPago & ","
                                StrSql = StrSql & "'" & Month(Fechadesde) & Year(Fechadesde) & "',"
                                StrSql = StrSql & "'" & Month(Fechahasta) & Year(Fechahasta) & "',"
                                StrSql = StrSql & "0," 'Renta imponible
                                StrSql = StrSql & "'" & Mid(arregloEstruc(11), 1, 50) & "'," 'Regimen Previsional
                                StrSql = StrSql & arregloEstruc(12) & "," 'Tipo Trabajador
                                StrSql = StrSql & "0," 'Dias Trabajados
                                StrSql = StrSql & "01" & "," 'Tipo de Linea
                                StrSql = StrSql & arregloMov(1) & "," 'Codigo Movimiento de personal
                                StrSql = StrSql & IIf(arregloFecD(1) = vbNull, "Null", ConvFecha(arregloFecD(aux))) & "," 'Fecha Desde para el movimiento
                                StrSql = StrSql & IIf(arregloFecH(1) = vbNull, "Null", ConvFecha(arregloFecH(aux))) & "," 'Fecha Hasta para el movimiento
                                StrSql = StrSql & "' '," 'Tramo Asignacion Familiar
                                StrSql = StrSql & "0," 'Numero de cargas simples
                                StrSql = StrSql & "0," 'Numero de cargas Maternales
                                StrSql = StrSql & "0," 'Numero de cargas Invalidas
                                StrSql = StrSql & "0," 'ASignacion Familiar
                                StrSql = StrSql & "0," 'ASignacion Familiar Retroactiva
                                StrSql = StrSql & "0," 'Renta Carga Familiares
                                StrSql = StrSql & "''," 'Solicitud Subsidio Trabajador Joven
                                'StrSql = StrSql & arregloEstruc(26) & "," 'Codigo AFP
                                If arreglo(26) = 0 Then
                                    StrSql = StrSql & arregloEstruc(26) & "," 'Codigo AFP
                                Else
                                    StrSql = StrSql & arreglo(26) & "," 'Codigo AFP
                                End If
                                StrSql = StrSql & "0," 'Renta Imponible AFP
                                StrSql = StrSql & "0," 'Cotizacion Obligatoria AFP
                                StrSql = StrSql & "0," 'Aporte Seguro Invalidez y Sobervivencia
                                StrSql = StrSql & "0," 'Cuenta de Ahorro Voluntaria
                                StrSql = StrSql & "0," 'Renta Imponible Sust a AFP
                                StrSql = StrSql & "0.00," 'Tasa Pactada
                                StrSql = StrSql & "0," 'Aporte Indem
                                StrSql = StrSql & "0," 'Numero de Periodos REVISAR!!!!!!!!!!!!!!!!!!!!!
                                StrSql = StrSql & "Null" & ","  'REVISAR CAMBIAR
                                StrSql = StrSql & "Null" & ","  'REVISAR CAMBIAR
                                StrSql = StrSql & "Null," 'Puesto de trabajo Pesado
                                StrSql = StrSql & "00.00," 'Porcentaje Cotizacion Trabajo Pesado
                                StrSql = StrSql & "0," 'Cotizacion Trabajo Pesado

                                StrSql = StrSql & arregloAPVI(aux).Cod & "," 'Inst Autor APVI
                                StrSql = StrSql & arregloEstruc(41) & "," 'Numero de contrato APVI
                                StrSql = StrSql & arregloEstruc(42) & "," 'Forma de Pago APVI
                                StrSql = StrSql & arregloAPVI(aux).Cotiza & "," 'Cotizacion APVI
                                StrSql = StrSql & arregloAPVI(aux).Depositos & "," 'Cotizacion Depositos convenidos
                                
                                If total_APVC > 0 Then
                                    StrSql = StrSql & arregloAPVC(1).Cod & ","  'Codigo Institucion Autorizada APVC
                                    StrSql = StrSql & arregloEstruc(46) & ","  'Numero de Contrato APVC
                                    StrSql = StrSql & arregloEstruc(47) & ","  'Forma de Pago APVC
                                    StrSql = StrSql & arregloAPVC(1).Cotiza & "," 'Cotizacion Trabajador APVC
                                    StrSql = StrSql & arregloAPVC(1).Depositos & "," 'Cotizacion Empleador APVC
                                Else
                                    StrSql = StrSql & "0,"  'Codigo Institucion Autorizada APVC
                                    StrSql = StrSql & "0,"  'Numero de Contrato APVC
                                    StrSql = StrSql & "0,"  'Forma de Pago APVC
                                    StrSql = StrSql & "0," 'Cotizacion Trabajador APVC
                                    StrSql = StrSql & "0," 'Cotizacion Empleador APVC
                                End If
                                
                                StrSql = StrSql & "'" & "" & "'," 'RUT Afiliado Voluntario
                                StrSql = StrSql & "'" & "" & "'," 'DV Afiliado Voluntario
                                StrSql = StrSql & "'" & "" & "'," 'Apellido Paterno Afiliado Voluntario
                                StrSql = StrSql & "'" & "" & "'," 'Apellido Materno Afiliado Voluntario
                                StrSql = StrSql & "'" & "" & "'," 'Nombres Afiliado Voluntario
                                StrSql = StrSql & "0," 'Codigo Movimiento Personal Afiliado Voluntario
                                StrSql = StrSql & "Null" & "," 'Fecha Desde Afiliado Voluntario
                                StrSql = StrSql & "Null" & "," 'Fecha Hasta Afiliado Voluntario
                                StrSql = StrSql & "0," 'Codigo de la AFP
                                StrSql = StrSql & "0," 'Monto Capitalizacion Voluntaria
                                StrSql = StrSql & "0," 'Monto Ahorro Voluntario
                                StrSql = StrSql & "0," 'Nu8mero de Periodos de Cotizacion
                                StrSql = StrSql & arregloEstruc(62) & "," 'Codigo Caja Regimen
                                StrSql = StrSql & "00.00," 'Tasa cotizacion EX cajas de Regimen
                                StrSql = StrSql & "0," 'Renta Imponible IPS
                                StrSql = StrSql & "0," 'Cotizacion Obligatoria INP
                                StrSql = StrSql & "0," 'Renta Imponible Desahucio
                                StrSql = StrSql & "0," 'Codigo ex caja Regimen
                                StrSql = StrSql & "0," 'Tasa Cotizacion Desahucio
                                StrSql = StrSql & "0," 'Cotizacion Desahucio
                                StrSql = StrSql & "0," 'Cotizacion Fonasa
                                StrSql = StrSql & "0," 'Cotizacion Accidente de Trabajo
                                StrSql = StrSql & "0," 'Bonificacion ley 15.386
                                StrSql = StrSql & "0," 'Descuento por Cargas Familiares
                                StrSql = StrSql & "0," 'Bonos Gobierno
                                StrSql = StrSql & arregloEstruc(75) & "," 'Codigo Institucion de Salud
                                StrSql = StrSql & "0,"  'FUN
                                StrSql = StrSql & "0," 'Renta Imponible Isapre
                                StrSql = StrSql & "0," 'Moneda del plan pactado con Isapre
                                StrSql = StrSql & "0," 'Cotizacion Pactada
                                StrSql = StrSql & "0," 'Cotizacion Obligatoria ISAPRE
                                StrSql = StrSql & "0," 'Cotizacion adicional Voluntaria
                                StrSql = StrSql & "0," 'Monto Garantia Explicito de Salud - GES
                                StrSql = StrSql & arregloEstruc(83) & "," 'Codigo CCAF
                                StrSql = StrSql & "0," 'Renta imponible CCAF
                                StrSql = StrSql & "0," 'Creditos Personales CCAF
                                StrSql = StrSql & "0," 'Descuento Dental
                                StrSql = StrSql & "0," 'Descuento por Leasing
                                StrSql = StrSql & "0," 'Descuentos por Seguro de Vida
                                StrSql = StrSql & "0," 'Otros Descuentos CCAF
                                StrSql = StrSql & "0," 'Cotiz CCAF de no Afiliados ISAPRE
                                StrSql = StrSql & "0," 'Descuentos Cargas Familiares CCAF
                                StrSql = StrSql & "0," 'Otros Descuentos CCAF 1
                                StrSql = StrSql & "0," 'Otros Descuentos CCAF 2
                                StrSql = StrSql & "0," 'Bonos Gobierno
                                StrSql = StrSql & arregloEstruc(95) & "," 'Codigo de Sucursal
                                StrSql = StrSql & arregloEstruc(96) & "," 'Codigo Mutual
                                StrSql = StrSql & "0," 'Renta Imponible Mutual
                                StrSql = StrSql & "0," 'Cotizacion Accidente del trabajo
                                StrSql = StrSql & arregloEstruc(99) & "," 'Sucursal Para Pago Mutual
                                StrSql = StrSql & "0," 'Renta total Imponible Seguro de Cesantia
                                StrSql = StrSql & "0," 'Aporte Trabajador Seguro de Cesantia
                                StrSql = StrSql & "0," 'Aporte Empleador Seguro de Cesantia
                                StrSql = StrSql & "Null" & "," 'Rut Pagadora Subsidio
                                StrSql = StrSql & "Null" & "," 'DV Pagadora Subsidio
                                StrSql = StrSql & "'" & arregloEstruc(105) & "',"   'Centro Costo, sucursal, etc
                                StrSql = StrSql & IIf(HuboError, -1, 0)
                                
                                StrSql = StrSql & ")"
                                Flog.writeline
                                Flog.writeline "Insertando 3 : " & StrSql
                                objConn.Execute StrSql, , adExecuteNoRecords
                                Flog.writeline
                                Flog.writeline
                                'Sumo el numero de linea
                                Num_linea = Num_linea + 1
                                Flog.writeline
                                Flog.writeline "SE GRABO LINEA ADICIONAL DE APVI"
                                Flog.writeline
            
                                aux = aux + 1
                            Loop
                            
                            'Aperturas de lineas de APVC
                            aux = 2
                            Do While aux <= total_APVC And (aux < 30)
                               'Inserto en rep_previred las lineas adicionales
                                StrSql = "INSERT INTO rep_previred (bpronro,ternro,num_linea,Titulo,pliqnro_Desde,pliqnro_hasta,empnro,rut,DV,Apellido,Apellido2,Nombres,sexo,nacionalidad,tipo_pago,Periodo_desde,Periodo_hasta,renta_imp,reg_pre,TipTrabajador,DiasTrab,TipoDeLinea,CodmovPer,fechadesde"
                                StrSql = StrSql & ",Fechahasta , TramoAsigFam, NumCargasSim, NumCargasMat, NumCargasInv, AsigFam, AsigFamRetro, ReintCarFam, SolicSubsidioTrabJoven, CodAFP, RentaImponibleAFP, CotizObligAFP, AporteSIS, CAVoluntaAFP, RenImpSustAFP, TasaPact, AportIndem, NumPeriodos, PeriDesdeAFP"
                                StrSql = StrSql & ",PeriHastaAFP,PuesTrabPesado,PorcCotizTrabPesa,CotizTrabPesa,InstAutAPV,NumContratoAPVI,ForPagAPV,CotizAPV,CotizDepConv,CodInstAutorizadaAPVC,NumContratoAPVC,FPagoAPVC,CotizTrabajadorAPVC,CotizEmpleadorAPVC,RUTAfVolunt,DVAfVolunt,ApePatVolunt"
                                StrSql = StrSql & ",ApeMatVolunt,NombVolunt,CodMovPersVolunt,FecDesdeVolunt,FecHastaVolunt,CodAFPVolunt,MontoCapVolunt,MontoAhorroVolunt,NumPerVolunt,CodCaReg,TasaCotCajPrev,RentaImpIPS,CotizObligINP,RentImpoDesah,CodCaRegDesah,TasaCotDesah,CotizDesah"
                                StrSql = StrSql & ",CotizFonasa,CotizAccTrab,BonLeyInp,DescCargFam,BonosGobierno,CodInstSal,NumFun,RentaImpIsapre,MonPlanIsapre,CotizPact,CotizObligIsapre,CotizAdicVolun,MontoGarantiaSaludGES,CodCCAF,RentaImponibleCCAF,CredPerCCAF,DescDentCCAF"
                                StrSql = StrSql & ",DescLeasCCAF,DescVidaCCAF,OtrosDesCCAF,CotCCAFnoIsapre,DesCarFamCCAF,OtrosDesCCAF1,OtrosDesCCAF2,BonosGobiernoCCAF,CodigoSucursalCCAF,CodMut,RentaimpMut,CotizAccTrabMut,SucPagMut,RentTotImp,AporTrabSeg,AporEmpSeg,RutPag,DVPag,CentroCosto,auxdeci"
                                
                                StrSql = StrSql & ") VALUES ("
                                
                                StrSql = StrSql & NroProcesoBatch & ","
                                StrSql = StrSql & rs_Empleados!ternro & ","
                                StrSql = StrSql & Num_linea & ","
                                StrSql = StrSql & "'" & Left(Titulo, 50) & "',"
                                StrSql = StrSql & "Null,"
                                StrSql = StrSql & "Null,"
                                StrSql = StrSql & Empresa & ","
                                StrSql = StrSql & "'" & RUT & "',"
                                StrSql = StrSql & "'" & DV & "',"
                                StrSql = StrSql & "'" & Apellido & "',"
                                StrSql = StrSql & "'" & Apellido2 & "',"
                                StrSql = StrSql & "'" & NombreEmp & "',"
                                StrSql = StrSql & "'" & Sexo & "',"
                                StrSql = StrSql & "'" & Nacionalidad & "',"
                                StrSql = StrSql & TipoPago & ","
                                StrSql = StrSql & "'" & Month(Fechadesde) & Year(Fechadesde) & "',"
                                StrSql = StrSql & "'" & Month(Fechahasta) & Year(Fechahasta) & "',"
                                StrSql = StrSql & "0," 'Renta imponible
                                StrSql = StrSql & "'" & Mid(arregloEstruc(11), 1, 50) & "'," 'Regimen Previsional
                                StrSql = StrSql & arregloEstruc(12) & "," 'Tipo Trabajador
                                StrSql = StrSql & "0," 'Dias Trabajados
                                StrSql = StrSql & "01" & "," 'Tipo de Linea
                                StrSql = StrSql & arregloMov(aux) & "," 'Codigo Movimiento de personal
                                StrSql = StrSql & IIf(arregloFecD(aux) = vbNull, "Null", ConvFecha(arregloFecD(aux))) & "," 'Fecha Desde para el movimiento
                                StrSql = StrSql & IIf(arregloFecH(aux) = vbNull, "Null", ConvFecha(arregloFecH(aux))) & "," 'Fecha Hasta para el movimiento
                                StrSql = StrSql & "' '," 'Tramo Asignacion Familiar
                                StrSql = StrSql & "0," 'Numero de cargas simples
                                StrSql = StrSql & "0," 'Numero de cargas Maternales
                                StrSql = StrSql & "0," 'Numero de cargas Invalidas
                                StrSql = StrSql & "0," 'ASignacion Familiar
                                StrSql = StrSql & "0," 'ASignacion Familiar Retroactiva
                                StrSql = StrSql & "0," 'Renta Carga Familiares
                                StrSql = StrSql & "''," 'Solicitud Subsidio Trabajador Joven
                                'StrSql = StrSql & arregloEstruc(26) & "," 'Codigo AFP
                                If arreglo(26) = 0 Then
                                   StrSql = StrSql & arregloEstruc(26) & "," 'Codigo AFP
                                Else
                                   StrSql = StrSql & arreglo(26) & "," 'Codigo AFP
                                End If
                                StrSql = StrSql & "0," 'Renta Imponible AFP
                                StrSql = StrSql & "0," 'Cotizacion Obligatoria AFP
                                StrSql = StrSql & "0," 'Aporte Seguro Invalidez y Sobervivencia
                                StrSql = StrSql & "0," 'Cuenta de Ahorro Voluntaria
                                StrSql = StrSql & "0," 'Renta Imponible Sust a AFP
                                StrSql = StrSql & "0.00," 'Tasa Pactada
                                StrSql = StrSql & "0," 'Aporte Indem
                                StrSql = StrSql & "0," 'Numero de Periodos REVISAR!!!!!!!!!!!!!!!!!!!!!
                                StrSql = StrSql & "Null" & ","  'REVISAR CAMBIAR
                                StrSql = StrSql & "Null" & ","  'REVISAR CAMBIAR
                                StrSql = StrSql & "Null," 'Puesto de trabajo Pesado
                                StrSql = StrSql & "00.00," 'Porcentaje Cotizacion Trabajo Pesado
                                StrSql = StrSql & "0," 'Cotizacion Trabajo Pesado
                                
                                If total_APVI > 0 Then
                                    StrSql = StrSql & arregloAPVI(1).Cod & "," 'Inst Autor APVI
                                    StrSql = StrSql & arregloEstruc(41) & "," 'Numero de contrato APVI
                                    StrSql = StrSql & arregloEstruc(42) & "," 'Forma de Pago APVI
                                    StrSql = StrSql & arregloAPVI(1).Cotiza & "," 'Cotizacion APVI
                                    StrSql = StrSql & arregloAPVI(1).Depositos & "," 'Cotizacion Depositos convenidos
                                Else
                                    StrSql = StrSql & "0," 'Inst Autor APVI
                                    StrSql = StrSql & "0," 'Numero de contrato APVI
                                    StrSql = StrSql & "0," 'Forma de Pago APVI
                                    StrSql = StrSql & "0," 'Cotizacion APVI
                                    StrSql = StrSql & "0," 'Cotizacion Depositos convenidos
                                End If
                                
                                StrSql = StrSql & arregloAPVC(aux).Cod & ","  'Codigo Institucion Autorizada APVC
                                StrSql = StrSql & arregloEstruc(46) & ","  'Numero de Contrato APVC
                                StrSql = StrSql & arregloEstruc(47) & ","  'Forma de Pago APVC
                                StrSql = StrSql & arregloAPVC(aux).Cotiza & "," 'Cotizacion Trabajador APVC
                                StrSql = StrSql & arregloAPVC(aux).Depositos & "," 'Cotizacion Empleador APVC
                                
                                StrSql = StrSql & "'" & "" & "'," 'RUT Afiliado Voluntario
                                StrSql = StrSql & "'" & "" & "'," 'DV Afiliado Voluntario
                                StrSql = StrSql & "'" & "" & "'," 'Apellido Paterno Afiliado Voluntario
                                StrSql = StrSql & "'" & "" & "'," 'Apellido Materno Afiliado Voluntario
                                StrSql = StrSql & "'" & "" & "'," 'Nombres Afiliado Voluntario
                                StrSql = StrSql & "0," 'Codigo Movimiento Personal Afiliado Voluntario
                                StrSql = StrSql & "Null" & "," 'Fecha Desde Afiliado Voluntario
                                StrSql = StrSql & "Null" & "," 'Fecha Hasta Afiliado Voluntario
                                StrSql = StrSql & "0," 'Codigo de la AFP
                                StrSql = StrSql & "0," 'Monto Capitalizacion Voluntaria
                                StrSql = StrSql & "0," 'Monto Ahorro Voluntario
                                StrSql = StrSql & "0," 'Nu8mero de Periodos de Cotizacion
                                StrSql = StrSql & arregloEstruc(62) & "," 'Codigo Caja Regimen
                                StrSql = StrSql & "00.00," 'Tasa cotizacion EX cajas de Regimen
                                StrSql = StrSql & "0," 'Renta Imponible IPS
                                StrSql = StrSql & "0," 'Cotizacion Obligatoria INP
                                StrSql = StrSql & "0," 'Renta Imponible Desahucio
                                StrSql = StrSql & "0," 'Codigo ex caja Regimen
                                StrSql = StrSql & "0," 'Tasa Cotizacion Desahucio
                                StrSql = StrSql & "0," 'Cotizacion Desahucio
                                StrSql = StrSql & "0," 'Cotizacion Fonasa
                                StrSql = StrSql & "0," 'Cotizacion Accidente de Trabajo
                                StrSql = StrSql & "0," 'Bonificacion ley 15.386
                                StrSql = StrSql & "0," 'Descuento por Cargas Familiares
                                StrSql = StrSql & "0," 'Bonos Gobierno
                                StrSql = StrSql & arregloEstruc(75) & "," 'Codigo Institucion de Salud
                                StrSql = StrSql & "0,"  'FUN
                                StrSql = StrSql & "0," 'Renta Imponible Isapre
                                StrSql = StrSql & "0," 'Moneda del plan pactado con Isapre
                                StrSql = StrSql & "0," 'Cotizacion Pactada
                                StrSql = StrSql & "0," 'Cotizacion Obligatoria ISAPRE
                                StrSql = StrSql & "0," 'Cotizacion adicional Voluntaria
                                StrSql = StrSql & "0," 'Monto Garantia Explicito de Salud - GES
                                StrSql = StrSql & arregloEstruc(83) & "," 'Codigo CCAF
                                StrSql = StrSql & "0," 'Renta imponible CCAF
                                StrSql = StrSql & "0," 'Creditos Personales CCAF
                                StrSql = StrSql & "0," 'Descuento Dental
                                StrSql = StrSql & "0," 'Descuento por Leasing
                                StrSql = StrSql & "0," 'Descuentos por Seguro de Vida
                                StrSql = StrSql & "0," 'Otros Descuentos CCAF
                                StrSql = StrSql & "0," 'Cotiz CCAF de no Afiliados ISAPRE
                                StrSql = StrSql & "0," 'Descuentos Cargas Familiares CCAF
                                StrSql = StrSql & "0," 'Otros Descuentos CCAF 1
                                StrSql = StrSql & "0," 'Otros Descuentos CCAF 2
                                StrSql = StrSql & "0," 'Bonos Gobierno
                                StrSql = StrSql & arregloEstruc(95) & "," 'Codigo de Sucursal
                                StrSql = StrSql & arregloEstruc(96) & "," 'Codigo Mutual
                                StrSql = StrSql & "0," 'Renta Imponible Mutual
                                StrSql = StrSql & "0," 'Cotizacion Accidente del trabajo
                                StrSql = StrSql & arregloEstruc(99) & "," 'Sucursal Para Pago Mutual
                                StrSql = StrSql & "0," 'Renta total Imponible Seguro de Cesantia
                                StrSql = StrSql & "0," 'Aporte Trabajador Seguro de Cesantia
                                StrSql = StrSql & "0," 'Aporte Empleador Seguro de Cesantia
                                StrSql = StrSql & "Null" & "," 'Rut Pagadora Subsidio
                                StrSql = StrSql & "Null" & "," 'DV Pagadora Subsidio
                                StrSql = StrSql & "'" & arregloEstruc(105) & "',"   'Centro Costo, sucursal, etc
                                StrSql = StrSql & IIf(HuboError, -1, 0)
                                
                                StrSql = StrSql & ")"
                                Flog.writeline
                                Flog.writeline "Insertando 4 : " & StrSql
                                objConn.Execute StrSql, , adExecuteNoRecords
                                Flog.writeline
                                Flog.writeline
                                'Sumo el numero de linea
                                Num_linea = Num_linea + 1
                                Flog.writeline
                                Flog.writeline "SE GRABO LINEA ADICIONAL POR APVC"
                                Flog.writeline
            
                                aux = aux + 1
                            Loop
                            
                        Case 2: 'Gratificaciones
                            
                            'Busco las fechas desde/hasta de los periodos de reliquidacion
                            
                            FechaDesdePeri = Fechadesde
                            FechaHastaPeri = Fechahasta
                            
                            StrSql = "SELECT " & TOP(1) & " pliqdesde FROM impuni_peri "
                            StrSql = StrSql & " inner join periodo on periodo.pliqnro = impuni_peri.pliqnro "
                            StrSql = StrSql & " Where pronro = " & rs_Empleados!pronro
                            StrSql = StrSql & " order by pliqdesde"
                            OpenRecordset StrSql, rs_impuniperi
                            If Not rs_impuniperi.EOF Then
                                    FechaDesdePeri = rs_impuniperi!pliqdesde
                            End If
                            
                            StrSql = "SELECT " & TOP(1) & " pliqhasta FROM impuni_peri "
                            StrSql = StrSql & " inner join periodo on periodo.pliqnro = impuni_peri.pliqnro "
                            StrSql = StrSql & " Where pronro = " & rs_Empleados!pronro
                            StrSql = StrSql & " order by pliqhasta desc"
                            OpenRecordset StrSql, rs_impuniperi
                            If Not rs_impuniperi.EOF Then
                                    FechaHastaPeri = rs_impuniperi!pliqhasta
                            End If
                            
                            
                            'Inserto en rep_previred
                                                        
                            StrSql = "INSERT INTO rep_previred (bpronro,ternro,num_linea,Titulo,pliqnro_Desde,pliqnro_hasta,empnro,rut,DV,Apellido,Apellido2,Nombres,sexo,nacionalidad,tipo_pago,Periodo_desde,Periodo_hasta,renta_imp,reg_pre,TipTrabajador,DiasTrab,TipoDeLinea,CodmovPer,fechadesde"
                            StrSql = StrSql & ",Fechahasta , TramoAsigFam, NumCargasSim, NumCargasMat, NumCargasInv, AsigFam, AsigFamRetro, ReintCarFam, SolicSubsidioTrabJoven, CodAFP, RentaImponibleAFP, CotizObligAFP, AporteSIS, CAVoluntaAFP, RenImpSustAFP, TasaPact, AportIndem, NumPeriodos, PeriDesdeAFP"
                            StrSql = StrSql & ",PeriHastaAFP,PuesTrabPesado,PorcCotizTrabPesa,CotizTrabPesa,InstAutAPV,NumContratoAPVI,ForPagAPV,CotizAPV,CotizDepConv,CodInstAutorizadaAPVC,NumContratoAPVC,FPagoAPVC,CotizTrabajadorAPVC,CotizEmpleadorAPVC,RUTAfVolunt,DVAfVolunt,ApePatVolunt"
                            StrSql = StrSql & ",ApeMatVolunt,NombVolunt,CodMovPersVolunt,FecDesdeVolunt,FecHastaVolunt,CodAFPVolunt,MontoCapVolunt,MontoAhorroVolunt,NumPerVolunt,CodCaReg,TasaCotCajPrev,RentaImpIPS,CotizObligINP,RentImpoDesah,CodCaRegDesah,TasaCotDesah,CotizDesah"
                            StrSql = StrSql & ",CotizFonasa,CotizAccTrab,BonLeyInp,DescCargFam,BonosGobierno,CodInstSal,NumFun,RentaImpIsapre,MonPlanIsapre,CotizPact,CotizObligIsapre,CotizAdicVolun,MontoGarantiaSaludGES,CodCCAF,RentaImponibleCCAF,CredPerCCAF,DescDentCCAF"
                            StrSql = StrSql & ",DescLeasCCAF,DescVidaCCAF,OtrosDesCCAF,CotCCAFnoIsapre,DesCarFamCCAF,OtrosDesCCAF1,OtrosDesCCAF2,BonosGobiernoCCAF,CodigoSucursalCCAF,CodMut,RentaimpMut,CotizAccTrabMut,SucPagMut,RentTotImp,AporTrabSeg,AporEmpSeg,RutPag,DVPag,CentroCosto,auxdeci"
                            
                            StrSql = StrSql & ") VALUES ("
                            
                            StrSql = StrSql & NroProcesoBatch & ","
                            StrSql = StrSql & rs_Empleados!ternro & ","
                            StrSql = StrSql & Num_linea & ","
                            StrSql = StrSql & "'" & Left(Titulo, 50) & "',"
                            StrSql = StrSql & "Null,"
                            StrSql = StrSql & "Null,"
                            StrSql = StrSql & Empresa & ","
                            StrSql = StrSql & "'" & RUT & "',"
                            StrSql = StrSql & "'" & DV & "',"
                            StrSql = StrSql & "'" & Apellido & "',"
                            StrSql = StrSql & "'" & Apellido2 & "',"
                            StrSql = StrSql & "'" & NombreEmp & "',"
                            StrSql = StrSql & "'" & Sexo & "',"
                            StrSql = StrSql & "'" & Nacionalidad & "',"
                            StrSql = StrSql & TipoPago & ","
                            StrSql = StrSql & "'" & Month(FechaDesdePeri) & Year(FechaDesdePeri) & "',"
                            StrSql = StrSql & "'" & Month(FechaHastaPeri) & Year(FechaHastaPeri) & "',"
                            StrSql = StrSql & arreglo(10) & "," 'Renta imponible
                            StrSql = StrSql & "'" & Mid(arregloEstruc(11), 1, 50) & "'," 'Regimen Previsional
                            StrSql = StrSql & arregloEstruc(12) & "," 'Tipo Trabajador
                            StrSql = StrSql & "30," 'Dias Trabajados
                            StrSql = StrSql & "'00'," 'Tipo de Linea
                            StrSql = StrSql & "0," 'Codigo Movimiento de personal
                            StrSql = StrSql & "null," 'Fecha Desde para el movimiento
                            StrSql = StrSql & "null," 'Fecha Hasta para el movimiento
                            StrSql = StrSql & "' '," 'Tramo Asignacion Familiar
                            StrSql = StrSql & "0," 'Numero de cargas simples
                            StrSql = StrSql & "0," 'Numero de cargas Maternales
                            StrSql = StrSql & "0," 'Numero de cargas Invalidas
                            StrSql = StrSql & "0," 'ASignacion Familiar
                            StrSql = StrSql & "0," 'ASignacion Familiar Retroactiva
                            StrSql = StrSql & "0," 'Renta Carga Familiares
                            StrSql = StrSql & "'N'," 'Solicitud Subsidio Trabajador Joven
                            'StrSql = StrSql & arregloEstruc(26) & "," 'Codigo AFP
                            If arreglo(26) = 0 Then
                                StrSql = StrSql & arregloEstruc(26) & "," 'Codigo AFP
                            Else
                                StrSql = StrSql & arreglo(26) & "," 'Codigo AFP
                            End If
                            StrSql = StrSql & arreglo(27) & ","  'Renta Imponible AFP
                            StrSql = StrSql & arreglo(28) & "," 'Cotizacion Obligatoria AFP
                            StrSql = StrSql & arreglo(29) & "," 'Aporte Seguro Invalidez y Sobervivencia
                            StrSql = StrSql & "0," 'Cuenta de Ahorro Voluntaria
                            StrSql = StrSql & "0," 'Renta Imponible Sust a AFP
                            StrSql = StrSql & "0.00," 'Tasa Pactada
                            StrSql = StrSql & "0," 'Aporte Indem
                            StrSql = StrSql & "0," 'Numero de Periodos REVISAR!!!!!!!!!!!!!!!!!!!!!
                            StrSql = StrSql & "Null" & ","  'Periodo desde rent imp sust
                            StrSql = StrSql & "Null" & ","  'Periodo hasta rent imp sust
                            StrSql = StrSql & "Null," 'Puesto de trabajo Pesado
                            StrSql = StrSql & "00.00," 'Porcentaje Cotizacion Trabajo Pesado
                            StrSql = StrSql & "0," 'Cotizacion Trabajo Pesado
                            StrSql = StrSql & "000," 'codigo  Inst  APVI
                            'StrSql = StrSql & "null," 'Numero de contrato APVI
                            StrSql = StrSql & "'" & "" & "'," 'Numero de contrato APVI
                            StrSql = StrSql & "0,"  'Forma de Pago APVI
                            StrSql = StrSql & "0," 'Cotizacion APVI
                            StrSql = StrSql & "0," 'Cotizacion Depositos convenidos
                            StrSql = StrSql & "000," 'Codigo Institucion Autorizada APVC
                            'StrSql = StrSql & "'0'," 'Numero de Contrato APVC
                            StrSql = StrSql & "'" & "" & "'," 'Numero de Contrato APVC
                            StrSql = StrSql & "0," 'Forma de Pago APVC
                            StrSql = StrSql & "0," 'Cotizacion Trabajador APVC
                            StrSql = StrSql & "0," 'Cotizacion Empleador APVC
                            'StrSql = StrSql & "'" & "" & "'," 'RUT Afiliado Voluntario
                            StrSql = StrSql & "'0'," 'RUT Afiliado Voluntario
                            StrSql = StrSql & "'" & "" & "'," 'DV Afiliado Voluntario
                            StrSql = StrSql & "'" & "" & "'," 'Apellido Paterno Afiliado Voluntario
                            StrSql = StrSql & "'" & "" & "'," 'Apellido Materno Afiliado Voluntario
                            StrSql = StrSql & "'" & "" & "'," 'Nombres Afiliado Voluntario
                            StrSql = StrSql & "0," 'Codigo Movimiento Personal Afiliado Voluntario
                            StrSql = StrSql & "Null" & "," 'Fecha Desde Afiliado Voluntario
                            StrSql = StrSql & "Null" & "," 'Fecha Hasta Afiliado Voluntario
                            StrSql = StrSql & "0," 'Codigo de la AFP
                            StrSql = StrSql & "0," 'Monto Capitalizacion Voluntaria
                            StrSql = StrSql & "0," 'Monto Ahorro Voluntario
                            StrSql = StrSql & "0," 'Nu8mero de Periodos de Cotizacion
                            StrSql = StrSql & arregloEstruc(62) & "," 'Codigo Caja Regimen
                            StrSql = StrSql & arreglo(63) & "," 'Tasa cotizacion EX cajas de Regimen
                            StrSql = StrSql & arreglo(64) & "," 'Renta Imponible IPS
                            StrSql = StrSql & arreglo(65) & "," 'Cotizacion Obligatoria INP
                            StrSql = StrSql & arreglo(66) & "," 'Renta Imponible Desahucio
                            StrSql = StrSql & arreglo(67) & "," 'Codigo ex caja Regimen
                            StrSql = StrSql & arreglo(68) & "," 'Tasa Cotizacion Desahucio
                            StrSql = StrSql & arreglo(69) & "," 'Cotizacion Desahucio
                            StrSql = StrSql & arreglo(70) & "," 'Cotizacion Fonasa
                            StrSql = StrSql & arreglo(71) & "," 'Cotizacion Accidente de Trabajo
                            StrSql = StrSql & "0," 'Bonificacion ley 15.386
                            StrSql = StrSql & "0," 'Descuento por Cargas Familiares
                            StrSql = StrSql & "0," 'Bonos Gobierno
                            StrSql = StrSql & arregloEstruc(75) & "," 'Codigo Institucion de Salud
                            StrSql = StrSql & "0,"  'FUN
                            StrSql = StrSql & arreglo(77) & "," 'Renta Imponible Isapre
                            StrSql = StrSql & "1," 'Moneda del plan pactado con Isapre
                            StrSql = StrSql & "0," 'Cotizacion Pactada
                            StrSql = StrSql & arreglo(80) & "," 'Cotizacion Obligatoria ISAPRE
                            StrSql = StrSql & "0," 'Cotizacion adicional Voluntaria
                            StrSql = StrSql & "0," 'Monto Garantia Explicito de Salud - GES
                            StrSql = StrSql & arregloEstruc(83) & "," 'Codigo CCAF
                            'StrSql = StrSql & "0," 'Renta imponible CCAF
                            StrSql = StrSql & arreglo(84) & "," 'Renta imponible CCAF
                            StrSql = StrSql & "0," 'Creditos Personales CCAF
                            StrSql = StrSql & "0," 'Descuento Dental
                            StrSql = StrSql & "0," 'Descuento por Leasing
                            StrSql = StrSql & "0," 'Descuentos por Seguro de Vida
                            StrSql = StrSql & "0," 'Otros Descuentos CCAF
                            StrSql = StrSql & arreglo(90) & "," 'Cotiz CCAF de no Afiliados ISAPRE
                            StrSql = StrSql & "0," 'Descuentos Cargas Familiares CCAF
                            StrSql = StrSql & "0," 'Otros Descuentos CCAF 1
                            StrSql = StrSql & "0," 'Otros Descuentos CCAF 2
                            StrSql = StrSql & "0," 'Bonos Gobierno
                            'StrSql = StrSql & arregloEstruc(95) & "," 'Codigo de Sucursal
                            StrSql = StrSql & "0," 'Codigo de Sucursal
                            StrSql = StrSql & arregloEstruc(96) & "," 'Codigo Mutual
                            'StrSql = StrSql & "0," 'Renta Imponible Mutual
                            StrSql = StrSql & arreglo(97) & "," 'Renta Imponible Mutual
                            StrSql = StrSql & arreglo(98) & "," 'Cotizacion Accidente del trabajo
                            'StrSql = StrSql & "0," 'Cotizacion Accidente del trabajo
                            StrSql = StrSql & arregloEstruc(99) & "," 'Sucursal Para Pago Mutual
                            StrSql = StrSql & IIf(EstaIPS And SeguroCesantia, 0, arreglo(100)) & "," 'Renta total Imponible Seguro de Cesantia
                            StrSql = StrSql & IIf(EstaIPS And SeguroCesantia, 0, arreglo(101)) & "," 'Aporte Trabajador Seguro de Cesantia
                            StrSql = StrSql & IIf(EstaIPS And SeguroCesantia, 0, arreglo(102)) & "," 'Aporte Empleador Seguro de Cesantia
                            StrSql = StrSql & "Null" & "," 'Rut Pagadora Subsidio
                            StrSql = StrSql & "Null" & "," 'DV Pagadora Subsidio
                            StrSql = StrSql & "'" & arregloEstruc(105) & "',"   'Centro Costo, sucursal, etc
                            StrSql = StrSql & IIf(HuboError, -1, 0)
                            
                            StrSql = StrSql & ")"
                            Flog.writeline
                            Flog.writeline "Insertando 5 : " & StrSql
                            objConn.Execute StrSql, , adExecuteNoRecords
                            Flog.writeline
                            Flog.writeline
                            'Sumo el numero de linea
                            Num_linea = Num_linea + 1
                            Flog.writeline
                            Flog.writeline "SE GRABO LINEA DE GRATIFICACION"
                            Flog.writeline
                    End Select
                    

                    
                'Else
                If HuboError Then
                    'Sumo 1 A la cantidad de errores
                    CantEmplError = CantEmplError + 1
                    Flog.writeline
                    Flog.writeline "SE DETECTARON ERRORES EN EL EMPLEADO "
                    Flog.writeline
                    Errores = True
                End If
                    
                'Actualizo el progreso
                Progreso = Progreso + IncPorc
                TiempoAcumulado = GetTickCount
                
                If Errores = False Then
                    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & _
                    ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & _
                    "' WHERE bpronro = " & NroProcesoBatch
                    objconnProgreso.Execute StrSql, , adExecuteNoRecords
                Else
                    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & _
                    ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & _
                    "',bprcestado = 'Error' WHERE bpronro = " & NroProcesoBatch
                    objconnProgreso.Execute StrSql, , adExecuteNoRecords
                End If
                
                ' ----------------------------------------------------------------
                
                End If
                
                MyCommitTrans
                'Paso al siguiente Empleado
                rs_Empleados.MoveNext
    Loop
End If 'If Not HuboError

If rs_Empleados.State = adStateOpen Then rs_Empleados.Close
If rs_CantEmpleados.State = adStateOpen Then rs_CantEmpleados.Close
If rs_Acu_liq.State = adStateOpen Then rs_Acu_liq.Close
If rs_Confrep.State = adStateOpen Then rs_Confrep.Close
If rs_Conceptos.State = adStateOpen Then rs_Conceptos.Close
If rs_Detliq.State = adStateOpen Then rs_Detliq.Close
If rs_Tercero.State = adStateOpen Then rs_Tercero.Close
If rs_Estructura.State = adStateOpen Then rs_Estructura.Close
If rs_Rut.State = adStateOpen Then rs_Rut.Close
If rs_Estr_cod.State = adStateOpen Then rs_Estr_cod.Close
If rs_impuniperi.State = adStateOpen Then rs_impuniperi.Close

Set rs_Empleados = Nothing
Set rs_CantEmpleados = Nothing
Set rs_Acu_liq = Nothing
Set rs_Confrep = Nothing
Set rs_Conceptos = Nothing
Set rs_Detliq = Nothing
Set rs_Tercero = Nothing
Set rs_Estructura = Nothing
Set rs_Rut = Nothing
Set rs_Estr_cod = Nothing
Set rs_impuniperi = Nothing

Exit Sub

CE:
    Flog.writeline "=================================================================="
    Flog.writeline "Error: " & Err.Description
    Flog.writeline "Ultimo SQl Ejecutado: " & StrSql
    Flog.writeline "=================================================================="
    MyRollbackTrans
    'MyBeginTrans
    Progreso = Progreso + IncPorc
    TiempoAcumulado = GetTickCount
    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & _
             ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & _
             "' WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    'MyCommitTrans
    
    HuboError = True
    Flog.writeline " Error: " & Err.Description

End Sub

Public Function EsElUltimoEmpleado(ByVal rs As ADODB.Recordset, ByVal Anterior As Long) As Boolean
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
