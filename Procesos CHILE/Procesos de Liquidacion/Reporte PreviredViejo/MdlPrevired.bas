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

Const Version = "1.10 BIS"
Const FechaVersion = "11/01/2009"
'Modificacion = MB - Error en Ccosto en gratificaciones faltaba una coma

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

    On Error Resume Next
    'Abro la conexion
    OpenConnection strconexion, objConn
    If Err.Number <> 0 Then
        Flog.writeline "Problemas en la conexion"
        Exit Sub
    End If
    OpenConnection strconexion, objconnProgreso
    If Err.Number <> 0 Then
        Flog.writeline "Problemas en la conexion"
        Exit Sub
    End If
    On Error GoTo 0
    
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
' --------------------------------------------------------------------------------------------


Dim Empresa As Long
Dim Lista_Pro As String
Dim Fechadesde As Date
Dim Fechahasta As Date
Dim topeArreglo      As Integer   'USAR ESTA VARIABLE PARA EL TOPE
Dim arreglo(80)      As Double
Dim arregloEstruc(80) As String
Dim tNomina As Integer

Dim arregloMov(30) As Integer
Dim arregloFecD(30) As Date
Dim arregloFecH(30) As Date
Dim total_mov As Integer
Dim aux As Integer

Dim I      As Integer
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
Dim EstaINP As Boolean
Dim MesDesdeRec As String
Dim AnioDesdeRec As String
Dim MesHastaRec As String
Dim AnioHastaRec As String

'recordsets
Dim rs_Empleados As New ADODB.Recordset
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
    Do While Not rs_Empleados.EOF
    
    MyBeginTrans
          rs_Confrep.MoveFirst
          
          If rs_Empleados!ternro <> UltimoEmpleado Then  'Es el primero
                    
               UltimoEmpleado = rs_Empleados!ternro
                Flog.writeline "_______________________________________________________________________"
                
                'Buscar el apellido y nombre
                    StrSql = "SELECT * FROM tercero WHERE ternro = " & rs_Empleados!ternro
                    OpenRecordset StrSql, rs_Tercero
                    If Not rs_Tercero.EOF Then
                        Apellido = Left(rs_Tercero!terape, 30)
                        Apellido2 = Left(rs_Tercero!terape2, 30)
                        NombreEmp = Left(rs_Tercero!ternom, 22) & " " & Left(rs_Tercero!ternom2, 6)
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
                    
                    For Contador = 1 To 80
                        If (Contador = 52) Then 'HJI Existen para esta columna valores validos con cero.
                           arreglo(Contador) = -1 'Todo el arreglo deberia ser inicializado en -1 HJI
                        Else
                           arreglo(Contador) = 0
                        End If
                    Next Contador
                    
                    For Contador = 1 To 80
                        arregloEstruc(Contador) = ""
                    Next Contador
                    
                    For Contador = 1 To 30
                        arregloMov(Contador) = 0
                    Next Contador
                    
                    For Contador = 1 To 30
                        arregloFecD(Contador) = vbNull
                    Next Contador
                    
                    For Contador = 1 To 30
                        arregloFecH(Contador) = vbNull
                    Next Contador
                    Flog.writeline "Inicializar Arreglos de totales"
          End If
                            
                Do While Not rs_Confrep.EOF
                    Flog.writeline "Columna " & rs_Confrep!confnrocol
                    Select Case UCase(rs_Confrep!conftipo)
                    Case "AC":
                        StrSql = "SELECT * FROM acu_liq WHERE cliqnro = " & rs_Empleados!cliqnro & _
                                 " AND acunro =" & rs_Confrep!confval
                        OpenRecordset StrSql, rs_Acu_liq
                        If Not rs_Acu_liq.EOF Then
                                arreglo(rs_Confrep!confnrocol) = arreglo(rs_Confrep!confnrocol) + Abs(rs_Acu_liq!almonto)
                        End If
                       
                    Case "CO":
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
                                    If (arreglo(rs_Confrep!confnrocol) = -1 And rs_Confrep!confnrocol = 52) Then
                                        arreglo(rs_Confrep!confnrocol) = Abs(rs_Detliq!dlimonto)
                                    Else
                                        arreglo(rs_Confrep!confnrocol) = arreglo(rs_Confrep!confnrocol) + Abs(rs_Detliq!dlimonto)
                                    End If
                                End If
                            End If
                        End If
                        
                    Case "PCO":
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
                            End If
                        End If
                    
                    Case "TE": 'tipo estructura
                        
                            StrSql = " SELECT estructura.estrnro, estructura.estrdabr, estructura.estrcodext, htetdesde, htethasta "
                            StrSql = StrSql & " FROM his_estructura "
                            StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro "
                            StrSql = StrSql & " WHERE ternro = " & rs_Empleados!ternro & " AND "
                            StrSql = StrSql & " his_estructura.tenro = " & rs_Confrep!confval & "And "
                            StrSql = StrSql & " (his_estructura.htetdesde <= " & ConvFecha(Fechahasta) & ") And "
                            StrSql = StrSql & " ((" & ConvFecha(Fechahasta) & " <= his_estructura.htethasta) or (his_estructura.htethasta is null))"
                            OpenRecordset StrSql, rs_Estructura
                            
                        
                            If Not rs_Estructura.EOF Then
                                
                                Flog.writeline Espacios(Tabulador * 2) & "Estructura: " & rs_Estructura!estrnro & " - " & rs_Estructura!estrdabr
                                If rs_Confrep!confnrocol = 33 Then
                                    arregloEstruc(33) = rs_Estructura!estrdabr
                                    Flog.writeline Espacios(Tabulador * 2) & " OK "
                                Else
                                    StrSql = "SELECT * FROM estr_cod WHERE estrnro =" & rs_Estructura!estrnro
                                    StrSql = StrSql & " AND tcodnro = 38"
                                    OpenRecordset StrSql, rs_Estr_cod
                                
                                    If Not rs_Estr_cod.EOF Then
                                        arregloEstruc(rs_Confrep!confnrocol) = IIf(EsNulo(rs_Estr_cod!nrocod), "", Left(CStr(rs_Estr_cod!nrocod), 6))
                                        Flog.writeline Espacios(Tabulador * 2) & " OK "
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

                    Case "TM": 'tipo estructura
                            
                            'Hacer case de tipo de Movimiento y generar el array con las fechas correspondientes
                            total_mov = 0
                            'ALTA
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
                            StrSql = "SELECT fases.caunro, fases.altfec, fases.bajfec FROM fases "
                            StrSql = StrSql & " WHERE fases.real = -1 "
                            StrSql = StrSql & " AND fases.bajfec >=" & ConvFecha(Fechadesde)
                            StrSql = StrSql & " AND fases.bajfec <= " & ConvFecha(Fechahasta)
                            StrSql = StrSql & " AND fases.empleado = " & rs_Empleados!ternro
                            OpenRecordset StrSql, rs_Fases
                            Do While Not rs_Fases.EOF
                                total_mov = total_mov + 1
                                'segun la causa ==> busco la estructura y el codigo asociado


                                StrSql = "SELECT * FROM estr_cod "
                                StrSql = StrSql & " INNER JOIN causa_sitrev ON causa_sitrev.estrnro = estr_cod.estrnro"
                                StrSql = StrSql & " WHERE tcodnro = 38"
                                StrSql = StrSql & " AND causa_sitrev.caunro = " & rs_Fases!caunro
                                OpenRecordset StrSql, rs_Estr_cod
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
                            StrSql = " SELECT emp_lic.elfechadesde, emp_lic.elfechahasta, emp_lic.tdnro FROM emp_lic "
                            StrSql = StrSql & " WHERE emp_lic.empleado= " & rs_Empleados!ternro
                            StrSql = StrSql & " AND ((emp_lic.elfechadesde <= " & ConvFecha(Fechadesde)
                            StrSql = StrSql & " AND emp_lic.elfechahasta >= " & ConvFecha(Fechadesde) & ")"
                            StrSql = StrSql & " OR (emp_lic.elfechadesde >=" & ConvFecha(Fechadesde) & " AND emp_lic.elfechahasta <=" & ConvFecha(Fechahasta) & "))"
                            OpenRecordset StrSql, rs_Aux
                            Do While Not rs_Aux.EOF
                                total_mov = total_mov + 1
                                StrSql = "SELECT * FROM estr_cod "
                                StrSql = StrSql & " INNER JOIN csijp_srtd ON estr_cod.estrnro = csijp_srtd.estrnro "
                                StrSql = StrSql & " WHERE tcodnro = 38"
                                StrSql = StrSql & " AND csijp_srtd.tdnro = " & rs_Aux!tdnro
                                OpenRecordset StrSql, rs_Estr_cod
                                If Not rs_Estr_cod.EOF Then
                                   arregloMov(total_mov) = IIf(EsNulo(rs_Estr_cod!nrocod), "", Left(CStr(rs_Estr_cod!nrocod), 6))
                                   Flog.writeline Espacios(Tabulador * 2) & " OK "
                                   Flog.writeline
                                Else
                                   Flog.writeline Espacios(Tabulador * 2) & "  No se encontró el codigo Interno para el Movimiento."
                                   Flog.writeline
                                End If

                                If rs_Aux!elfechadesde < Fechadesde Then
                                    arregloFecD(total_mov) = Fechadesde
                                Else
                                    arregloFecD(total_mov) = rs_Aux!elfechadesde
                                End If
                                If rs_Aux!elfechahasta > Fechahasta Then
                                    arregloFecH(total_mov) = Fechahasta
                                Else
                                    arregloFecH(total_mov) = rs_Aux!elfechahasta
                                End If

                                rs_Aux.MoveNext
                            Loop
                            If total_mov = 0 Then
                                Flog.writeline Espacios(Tabulador * 2) & "Tipo de Estructura: " & rs_Confrep!confval
                                Flog.writeline Espacios(Tabulador * 2) & "No se encontraron movimientos."
                                Flog.writeline
                            Else
                                Flog.writeline Espacios(Tabulador * 2) & "se encontraron " & total_mov & " movimientos."
                                Flog.writeline
                            End If


                            '----------------------------------------------------------------------------------
                            ' FGZ - 20/11/2008 - se cambió todo esto por lo anterior
'                            StrSql = " SELECT estructura.estrnro, estructura.estrdabr, estructura.estrcodext, htetdesde, htethasta "
'                            StrSql = StrSql & " FROM his_estructura "
'                            StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro "
'                            StrSql = StrSql & " WHERE ternro = " & rs_Empleados!ternro & " AND "
'                            StrSql = StrSql & " his_estructura.tenro = " & rs_Confrep!confval & "And "
'                            StrSql = StrSql & " (his_estructura.htetdesde >= " & ConvFecha(Fechadesde) & ") And "
'                            StrSql = StrSql & " (" & ConvFecha(Fechahasta) & " >= his_estructura.htethasta)"
'                            OpenRecordset StrSql, rs_Estructura
'
'                            total_mov = 0
'
'                            Do While Not rs_Estructura.EOF
'
'                               total_mov = total_mov + 1
'                               Flog.writeline Espacios(Tabulador * 2) & "Estructura: " & rs_Estructura!Estrnro & " - " & rs_Estructura!estrdabr
'                               StrSql = "SELECT * FROM estr_cod WHERE estrnro =" & rs_Estructura!Estrnro
'                               StrSql = StrSql & " AND tcodnro = 38"
'                               OpenRecordset StrSql, rs_Estr_cod
'
'                               If Not rs_Estr_cod.EOF Then
'                                  arregloMov(total_mov) = IIf(EsNulo(rs_Estr_cod!nrocod), "", Left(CStr(rs_Estr_cod!nrocod), 6))
'                                  Flog.writeline Espacios(Tabulador * 2) & " OK "
'                                  Flog.writeline
'                               Else
'                                  Flog.writeline Espacios(Tabulador * 2) & "  No se encontró el codigo Interno para el Movimiento."
'                                  Flog.writeline
'                               End If
'                               rs_Estr_cod.Close
'
'                               If Not EsNulo(rs_Estructura!htetdesde) Then
'                                  arregloFecD(total_mov) = rs_Estructura!htetdesde
'                               End If
'
'                               If Not EsNulo(rs_Estructura!htethasta) Then
'                                  arregloFecH(total_mov) = rs_Estructura!htethasta
'                               End If
'
'                               rs_Estructura.MoveNext
'
'                            Loop
'                            If total_mov = 0 Then
'                                Flog.writeline Espacios(Tabulador * 2) & "Tipo de Estructura: " & rs_Confrep!confval
'                                Flog.writeline Espacios(Tabulador * 2) & "No se encontraron movimientos."
'                                Flog.writeline
'                            End If
'                            rs_Estructura.Close
                            
                            '----------------------------------------------------------------------------------
                            ' FGZ - 20/11/2008 - se cambió todo esto por lo anterior
                    Case "CTE": 'Constante
                                If rs_Confrep!confval2 = "" Or EsNulo(rs_Confrep!confval2) Then
                                    'Numerica
                                    arreglo(rs_Confrep!confnrocol) = rs_Confrep!confval
                                Else
                                    'Alfanumerica
                                    arregloEstruc(rs_Confrep!confnrocol) = rs_Confrep!confval2
                                End If
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
                    
                    For Contador = 1 To 80
                        If arregloEstruc(Contador) = "" Then
                            arregloEstruc(Contador) = "0"
                        End If
                    Next
                                    
                    '-----------------------------------------------------------------------------------
                                        
                                        
                    ' ----------------------------------------------------------------
                    ' Buscar el Rut DEL EMPLEADO
                    Flog.writeline
                    Flog.writeline "Procesando Campo 1 y 2. RUT y DV:  "
                    StrSql = " SELECT nrodoc FROM tercero " & _
                             " INNER JOIN ter_doc  ON (tercero.ternro = ter_doc.ternro AND ter_doc.tidnro = 1) " & _
                             " WHERE tercero.ternro= " & rs_Empleados!ternro
                    OpenRecordset StrSql, rs_Rut
          
                    If Not rs_Rut.EOF Then
                        RUT = Mid(rs_Rut!nrodoc, 1, Len(rs_Rut!nrodoc) - 1)
                        RUT = Replace(RUT, "-", "")
                        DV = Right(rs_Rut!nrodoc, 1)
                        
                        'HACER VALIDACION DE RUT Y DV
                        Flog.writeline Espacios(Tabulador * 1) & "RUT y DV Validos"
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
                    'Busco el valor de Tipo Pago
                    Flog.writeline "Procesando Campo 7: Tipo Pago "
'                        If rs_Empleados!ajugcias = -1 Then
'                            TipoPago = 2
'                        Else
'                            TipoPago = 1
'                        End If
                         TipoPago = tNomina
                    Flog.writeline Espacios(Tabulador * 1) & " Tipo Pago Obtenido "
                    Flog.writeline
                  '----------------------------------------------------------------
                                        
                  '-----------------------------------------------------------------------------------
                  ' Empieza Validaciones
                    
                    'Periodo de Remuneraciones   DESDE  Formato MMAAAA
                    'Este fue pasado por parametro
                    Flog.writeline "Procesando Campo 8: Periodo Remuneraciones Desde"
                    
                    Flog.writeline Espacios(Tabulador * 1) & "Periodo Remuneraciones Desde Obtenido "
                    Flog.writeline
                    'Periodo de Remuneraciones  HASTA   Formato MMAAAA
                    'Este fue pasado por parametro
                    Flog.writeline "Procesando Campo 9: Periodo Remuneraciones Hasta"
                    
                    Flog.writeline Espacios(Tabulador * 1) & "Periodo Remuneraciones Hasta Obtenido "
                    Flog.writeline
                    
                    
                    'Renta Imponible
                    Flog.writeline "Procesando Campo 10: Renta Imponible"
                        
                        
                        If arreglo(10) < 0 Then
                            Flog.writeline Espacios(Tabulador * 1) & "ERROR la Renta imponible no puede ser negativa"
                            HuboError = True
                        Else
                            Flog.writeline Espacios(Tabulador * 1) & "Renta Imponible Obtenida "
                        End If
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
                    
                    'Codigo de Movimiento del Personal
                    'PERMITE 0   08-03-07
                    Flog.writeline "Procesando Campo 14: Codigo de Movimiento del Personal"
                        
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
                    
                    'Fecha Desde
                    'Si Movimiento de personal es 1,3,4,5,6,7,8 el campo es obligatorio. QUEDA PARA REVISAR
                    Flog.writeline "Procesando Campo 15: Fecha Desde"
                    If arregloMov(aux) = "1" Or _
                       arregloMov(aux) = "3" Or _
                       arregloMov(aux) = "4" Or _
                       arregloMov(aux) = "5" Or _
                       arregloMov(aux) = "6" Or _
                       arregloMov(aux) = "7" Or _
                       arregloMov(aux) = "8" Then
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
                    Flog.writeline "Procesando Campo 16: Fecha Hasta"
                    If arregloMov(aux) = "2" Or _
                       arregloMov(aux) = "3" Or _
                       arregloMov(aux) = "4" Or _
                       arregloMov(aux) = "6" Then
                     Flog.writeline "Procesando Campo 16: Campo Obligatorio."
                     Flog.writeline Espacios(Tabulador * 1) & arregloFecH(aux)
                    Else
                     Flog.writeline Espacios(Tabulador * 1) & " Campo Optativo. "
                     arregloFecH(aux) = vbNull
                    End If
                    Flog.writeline

                    aux = aux + 1

                    Loop

                    'Tramo de Asignacion Familiar
                    Flog.writeline "Procesando Campo 17: Tramo de Asignacion Familiar"
'                    If arregloEstruc(17) = "0" Then
                    If arreglo(17) = 0 And arregloEstruc(17) = "0" Then
                        Flog.writeline Espacios(Tabulador * 1) & "Obligatorio. ERROR No se encontro el codigo para el  Tramo de Asignacion Familiar"
                        HuboError = True
                    Else
                        If arreglo(17) <> 0 Then
                            Select Case arreglo(17)
                                Case 1:
                                    arregloEstruc(17) = "A"
                                Case 2:
                                    arregloEstruc(17) = "B"
                                Case 3:
                                    arregloEstruc(17) = "C"
                                Case 4:
                                    arregloEstruc(17) = "D"
                                Case Else
                                    arregloEstruc(17) = "Sin"
                            End Select
                        End If
                        Flog.writeline Espacios(Tabulador * 1) & "Campo Obtenido"
                    End If
                    Flog.writeline
                    
                    'Num Cargas Simples
                    Flog.writeline "Procesando Campo 18: Numero de Cargas Simples"
                    
                    If arreglo(18) = 0 Then
                        Flog.writeline Espacios(Tabulador * 1) & "Campo Vacio. OPTATIVO..CONTINUA PROCESAMIENTO"
                    Else
                        If arreglo(18) < 0 Or arreglo(18) > 9 Then
                            Flog.writeline Espacios(Tabulador * 1) & "ERROR El Numero de Cargas Simples debe ser positivos y menores a 9"
                            HuboError = True
                        Else
                            Flog.writeline Espacios(Tabulador * 1) & "Numero de Cargas Simples Obtenido "
                        End If
                    End If
                    Flog.writeline
                    
                    'Num Cargas Maternales
                    Flog.writeline "Procesando Campo 19: Numero de Cargas Maternales"
                    
                    If arreglo(19) = 0 Then
                        Flog.writeline Espacios(Tabulador * 1) & "Campo Vacio. OPTATIVO..CONTINUA PROCESAMIENTO"
                    Else
                        If arreglo(19) < 0 Or arreglo(19) > 9 Then
                            Flog.writeline Espacios(Tabulador * 1) & "ERROR El Numero de Cargas Maternales debe ser positivos y menores a 9"
                            HuboError = True
                        Else
                            Flog.writeline Espacios(Tabulador * 1) & "Numero de Cargas Maternales Obtenido"
                        End If
                    End If
                    Flog.writeline
                    
                    
                    'Num Cargas Invalidas
                    Flog.writeline "Procesando Campo 20: Numero de Cargas Invalidas"
                    
                    If arreglo(20) = 0 Then
                        Flog.writeline Espacios(Tabulador * 1) & " Campo Vacio. OPTATIVO..CONTINUA PROCESAMIENTO"
                    Else
                        If arreglo(20) < 0 Or arreglo(20) > 9 Then
                            Flog.writeline Espacios(Tabulador * 1) & "ERROR El Numero de Cargas Invalidas debe ser positivos y menores a 9"
                            HuboError = True
                        Else
                            Flog.writeline Espacios(Tabulador * 1) & "Numero de Cargas Invalidas Obtenido"
                        End If
                    End If
                    Flog.writeline
                    
                    'Asignacion Familiar
                    Flog.writeline "Procesando Campo 21: Asignacion Familiar"
                    
                    If arreglo(21) = 0 Then
                        Flog.writeline Espacios(Tabulador * 1) & "Campo Vacio. OPTATIVO..CONTINUA PROCESAMIENTO"
                    Else
                        If arreglo(21) < 0 Or arreglo(21) > 600000 Then
                            Flog.writeline Espacios(Tabulador * 1) & "ERROR Asignacion Familiar debe ser positiva y menores a 600.000"
                            HuboError = True
                        Else
                            Flog.writeline Espacios(Tabulador * 1) & "Asignacion Familiar Obtenida"
                        End If
                    End If
                    Flog.writeline
                    
                    'Asignacion Retroactiva
                    Flog.writeline "Procesando Campo 22: Asignacion Retroactiva"
                    
                    If arreglo(22) = 0 Then
                        Flog.writeline Espacios(Tabulador * 1) & " Campo Vacio. OPTATIVO..CONTINUA PROCESAMIENTO"
                    Else
                        If arreglo(22) < 0 Or arreglo(22) > 600000 Then
                            Flog.writeline Espacios(Tabulador * 1) & "ERROR Asignacion Retroactiva debe ser positiva y menores a 600.000"
                            HuboError = True
                        Else
                            Flog.writeline Espacios(Tabulador * 1) & "Asignacion Retroactiva Obtenida"
                        End If
                    End If
                    Flog.writeline
                    
                    'Reintegro Cargas
                    Flog.writeline "Procesando Campo 23: Reintegro Cargas"
                    
                    If arreglo(23) = 0 Then
                        Flog.writeline Espacios(Tabulador * 1) & "Campo Vacio. OPTATIVO..CONTINUA PROCESAMIENTO"
                    Else
                        If arreglo(23) < 0 Or arreglo(23) > 600000 Then
                            Flog.writeline Espacios(Tabulador * 1) & "ERROR Reintegro Cargas debe ser positivo y menores a 600.000"
                            HuboError = True
                        Else
                            Flog.writeline Espacios(Tabulador * 1) & "Reintegro Cargas Obtenida"
                        End If
                    End If
                    Flog.writeline
                    
                    'Codigo de la AFP
                    Flog.writeline "Procesando Campo 24: Codigo de la AFP"
                    
                    If arregloEstruc(24) = "0" Then
                        Flog.writeline Espacios(Tabulador * 1) & "ERROR No se encontro el codigo para la AFP"
                        HuboError = True
                    Else
                        If IsNumeric(arregloEstruc(24)) = True Then
                            Flog.writeline Espacios(Tabulador * 1) & "Codigo de la AFP Obtenido "
                        Else
                            Flog.writeline Espacios(Tabulador * 1) & "ERROR. El campo contiene valores alfanuméricos y se requieren valores numéricos"
                            arregloEstruc(24) = "0"
                            HuboError = True
                        End If

                    End If
                    Flog.writeline
                    
                    'Cotizacion Obligatoria AFP
                    Flog.writeline "Procesando Campo 25: Cotizacion Obligatoria AFP"
                    'If arregloEstruc(12) = "0" And arreglo(13) > 0 Then Usar para campo 25
                    If arreglo(25) < 0 Then
                        Flog.writeline Espacios(Tabulador * 1) & "ERROR la cotizacion no puede ser negativa"
                        HuboError = True
                    Else
                        Flog.writeline Espacios(Tabulador * 1) & "Cotizacion Obligatoria AFP Obtenido "
                    End If
                  
                    Flog.writeline
                    
                    'Cuenta de Ahorro Voluntario AFP
                    Flog.writeline "Procesando Campo 26: Cuenta de Ahorro Voluntario AFP"
                    
                    If arreglo(26) = 0 Then
                        Flog.writeline Espacios(Tabulador * 1) & "Campo Vacio. OPTATIVO..CONTINUA PROCESAMIENTO"
                    Else
                        If arreglo(26) < 0 Then
                            Flog.writeline Espacios(Tabulador * 1) & "ERROR Cuenta de Ahorro Voluntario AFP no puede ser negativa"
                            HuboError = True
                        Else
                            Flog.writeline Espacios(Tabulador * 1) & "Cuenta de Ahorro Voluntario AFP Obtenido "
                        End If
                    End If
                    Flog.writeline
                    
                    'Renta Imp. Sustitutiva AFP
                    Flog.writeline "Procesando Campo 27: Renta Imp. Sustitutiva AFP"
                    
                    If arreglo(27) = 0 Then
                        Flog.writeline Espacios(Tabulador * 1) & " Campo Vacio. OPTATIVO..CONTINUA PROCESAMIENTO"
                    Else
                        If arreglo(27) < 0 Then
                            Flog.writeline Espacios(Tabulador * 1) & "ERROR Renta Imp. Sustitutiva AFP no puede ser negativo"
                            HuboError = True
                        Else
                            Flog.writeline Espacios(Tabulador * 1) & "Renta Imp. Sustitutiva AFP Obtenida "
                        End If
                    End If
                    Flog.writeline
                    
                    'Tasa Pactada
                    Flog.writeline "Procesando Campo 28: Tasa Pactada"
                    
                    If arreglo(28) = 0 Then
                        Flog.writeline Espacios(Tabulador * 1) & "Campo Vacio. OPTATIVO..CONTINUA PROCESAMIENTO"
                    Else
                        If arreglo(28) < 0 Then
                            Flog.writeline Espacios(Tabulador * 1) & "ERROR Tasa Pactada no puede ser negativo"
                            HuboError = True
                        Else
                            Flog.writeline Espacios(Tabulador * 1) & "Tasa Pactada Obtenida "
                        End If
                    End If
                    Flog.writeline
                    
                    'Aporte Indemnizacion Sustitutiva
                    Flog.writeline "Procesando Campo 29: Aporte Indemnizacion Sustitutiva"
                    
                    If arreglo(29) = 0 Then
                        Flog.writeline Espacios(Tabulador * 1) & "Campo Vacio. OPTATIVO..CONTINUA PROCESAMIENTO"
                    Else
                        If arreglo(29) < 0 Then
                            Flog.writeline Espacios(Tabulador * 1) & "ERROR Aporte Indemnizacion Sustitutiva no puede ser negativa"
                            HuboError = True
                        Else
                            Flog.writeline Espacios(Tabulador * 1) & "Aporte Indemnizacion Sustitutiva Obtenido "
                        End If
                    End If
                    Flog.writeline
                    
                    'Numero de Periodos Sustitutiva
                    Flog.writeline "Procesando Campo 30: Numero de Periodos Sustitutiva"
                    Flog.writeline Espacios(Tabulador * 1) & "Campo Optativo. NO DEFINIDO. Se pone a cero"
                    Flog.writeline
                    'Periodos Desde Sustitutiva
                    Flog.writeline "Procesando Campo 31: Periodos Desde Sustitutiva"
                    Flog.writeline Espacios(Tabulador * 1) & "Campo Optativo. NO DEFINIDO. Se pone a nulo"
                    Flog.writeline
                    'Periodos Hasta Sustitutiva
                    Flog.writeline "Procesando Campo 32: Periodos Hasta Sustitutiva"
                    Flog.writeline Espacios(Tabulador * 1) & "Campo Optativo. NO DEFINIDO. Se pone a nulo"
                    Flog.writeline
                    
                    'Puesto de Trabajo Pesado
                    Flog.writeline "Procesando Campo 33: Puesto de Trabajo Pesado"
                    If arregloEstruc(33) = "0" Then
                        Flog.writeline Espacios(Tabulador * 1) & "Campo Vacio. OPTATIVO..CONTINUA PROCESAMIENTO"
                        arregloEstruc(33) = ""
                    Else
                        Flog.writeline Espacios(Tabulador * 1) & "Puesto de Trabajo Pesado Obtenido "
                    End If
                    
                    Flog.writeline
                    
                    '% Cotizacion Trabajo Pesado
                    Flog.writeline "Procesando Campo 34: % Cotizacion Trabajo Pesado"
                    
                    If arreglo(34) < 0 Then
                        Flog.writeline Espacios(Tabulador * 1) & "ERROR el porcentaje cotizacion no puede ser negativo"
                        HuboError = True
                    Else
                        Flog.writeline Espacios(Tabulador * 1) & "% Cotizacion Trabajo Pesado Obtenida "
                    End If
                    Flog.writeline
                    
                    'Cotizacion Trabajo Pesado
                    Flog.writeline "Procesando Campo 35:  Cotizacion Trabajo Pesado"
                    
                    If arreglo(35) < 0 Then
                        Flog.writeline Espacios(Tabulador * 1) & "ERROR la cotizacion no puede ser negativa"
                        HuboError = True
                    Else
                        Flog.writeline Espacios(Tabulador * 1) & "Cotizacion Trabajo Pesado Obtenida "
                    End If
                    
                    Flog.writeline
                    
                    'Busco el RUT y el DV de la empresa
                    Flog.writeline "Procesando Campo 36:  RUT"
                    Flog.writeline "Procesando Campo 37:  DV"
                    Flog.writeline Espacios(Tabulador * 1) & "Campos Optativos. NO DEFINIDOS.Se ponen a nulo"
                    Flog.writeline
                    
                    'Codigo EX caja Regimen
                    Flog.writeline "Procesando Campo 38: Codigo EX caja Regimen"
                    
                    If arregloEstruc(38) = "0" Then
                        Flog.writeline Espacios(Tabulador * 1) & "Campo Vacio. OPTATIVO..CONTINUA PROCESAMIENTO"
                    Else
                        If IsNumeric(arregloEstruc(38)) = True Then
                            Flog.writeline Espacios(Tabulador * 1) & "Codigo EX caja Regimen Obtenido "
                        Else
                            Flog.writeline Espacios(Tabulador * 1) & "ERROR. El campo contiene valores alfanuméricos y se requieren valores numéricos"
                            arregloEstruc(38) = "0"
                            HuboError = True
                        End If
                        
                    End If
                    
                    Flog.writeline
                    
                    'Tasa Cotizacion Ex caja de Prevision
                    Flog.writeline "Procesando Campo 39:  Tasa Cotizacion Ex caja de Prevision"
                    
                    If arreglo(39) < 0 Then
                        Flog.writeline Espacios(Tabulador * 1) & "ERROR la Tasa de cotizacion de la Ex caja de Prevision no puede ser negativa"
                        HuboError = True
                    Else
                        Flog.writeline Espacios(Tabulador * 1) & "Tasa Cotizacion Ex caja Obtenida "
                    End If
                    
                    Flog.writeline
                   
                   'Cotizacion Obligatoria INP
                    
                    Flog.writeline "Procesando Campo 40:  Cotizacion Obligatoria INP"
                    If arreglo(40) < 0 Then
                        Flog.writeline Espacios(Tabulador * 1) & "ERROR la cotizacion INP no puede ser negativa"
                        HuboError = True
                    End If
                    
                    EstaINP = (UCase(arregloEstruc(11)) = "INP")
                    
                    ' Si no es negativa pregunto por la otra condicion de Validacion
                    If (arregloEstruc(12) = "0" And arreglo(13) > 0) And EstaINP Then
                             
                        If arreglo(40) = 0 Then
                            Flog.writeline Espacios(Tabulador * 1) & "ERROR la cotizacion INP no puede ser CERO ya que el trabajador es activo, dias trabajados es mayor a CERO y Regimen Previsional es INP"
                        Else
                            Flog.writeline Espacios(Tabulador * 1) & "Cotizazion Obligatoria Obtenida"
                        End If
                    Else
                        Flog.writeline Espacios(Tabulador * 1) & "Cotizazion Obligatoria Obtenida"
                    End If
                    Flog.writeline
                    
                    'Renta Imponible Desahucio
                    Flog.writeline "Procesando Campo 41:  Renta Imponible Desahucio"
                    
                    If arreglo(41) < 0 Then
                        Flog.writeline Espacios(Tabulador * 1) & "ERROR  Renta Imponible Desahucio no puede ser negativa"
                        HuboError = True
                    Else
                        If arreglo(41) > 60 Then
                            Flog.writeline Espacios(Tabulador * 1) & "ERROR la Renta Imponible Desahucio no puede ser mayor a 60"
                            HuboError = True
                        Else
                            If arreglo(41) = 0 Then
                                Flog.writeline Espacios(Tabulador * 1) & "Campo Vacio. OPTATIVO..CONTINUA PROCESAMIENTO"
                            Else
                                Flog.writeline Espacios(Tabulador * 1) & "Renta Imponible Desahucio Obtenida "
                            End If
                        End If
                    End If
                    
                    Flog.writeline
                    
                    'Codigo Ex caja Regimen Regimen Desahucio
                    Flog.writeline "Procesando Campo 42:  Codigo Ex caja Regimen Regimen Desahucio"
                    
                    If arregloEstruc(42) = "0" Then
                        Flog.writeline Espacios(Tabulador * 1) & "Campo Vacio. OPTATIVO..CONTINUA PROCESAMIENTO"
                    Else
                    
                        If IsNumeric(arregloEstruc(42)) = True Then
                            Flog.writeline Espacios(Tabulador * 1) & "Codigo EX caja Regimen Obtenido "
                        Else
                            Flog.writeline Espacios(Tabulador * 1) & "ERROR. El campo contiene valores alfanuméricos y se requieren valores numéricos"
                            arregloEstruc(42) = "0"
                            HuboError = True
                        End If
                        
                    End If
                    Flog.writeline
                    
                    'TASA Ex caja Regimen Regimen Desahucio
                    Flog.writeline "Procesando Campo 43:  Tasa caja Regimen Regimen Desahucio"
                    If arreglo(43) = 0 Then
                        Flog.writeline Espacios(Tabulador * 1) & "Campo Vacio. OPTATIVO..CONTINUA PROCESAMIENTO"
                    Else
                        Flog.writeline Espacios(Tabulador * 1) & "Tasa EX caja Regimen Obtenido "
                    End If
                    
                    Flog.writeline
                    
                    'Cotizacion Desahucio
                    Flog.writeline "Procesando Campo 44:  Cotizacion Desahucio"
                    
                    If arreglo(44) < 0 Then
                        Flog.writeline Espacios(Tabulador * 1) & "ERROR la Tasa de cotizacion no puede ser negativa"
                        HuboError = True
                    End If
                    
                    If arreglo(44) = 0 Then
                        Flog.writeline Espacios(Tabulador * 1) & "Campo Vacio. OPTATIVO..CONTINUA PROCESAMIENTO"
                    Else
                        Flog.writeline Espacios(Tabulador * 1) & "Cotizacion Desahucio Obtenida"
                    End If
                    Flog.writeline
                    
                    'Otros Aportes INP
                    Flog.writeline "Procesando Campo 45:  Otros Aportes INP"
                    Flog.writeline Espacios(Tabulador * 1) & "Campo Fijo. No Configurable."
                    Flog.writeline
                    
                    'Cotizacion Fonasa
                    Flog.writeline "Procesando Campo 46: Cotizacion Fonasa"
                    
                    If arreglo(46) < 0 Then
                        Flog.writeline Espacios(Tabulador * 1) & "ERROR la Tasa de cotizacion FONASA no puede ser negativa"
                        HuboError = True
                    End If
                    
                    If arreglo(46) = 0 Then
                        Flog.writeline Espacios(Tabulador * 1) & "Campo Vacio. OPTATIVO..CONTINUA PROCESAMIENTO"
                    Else
                        Flog.writeline Espacios(Tabulador * 1) & "Cotizacion Fonasa Obtenida"
                    End If
                    Flog.writeline
                    
                    'Cotizacion Acc de Trabajo
                    Flog.writeline "Procesando Campo 47: Cotizacion Acc de Trabajo"
                    
                    If arreglo(47) < 0 Then
                        Flog.writeline Espacios(Tabulador * 1) & "ERROR la Cotizacion Acc de Trabajo no puede ser negativa"
                        HuboError = True
                    End If
                    
                    If arreglo(47) = 0 Then
                        Flog.writeline Espacios(Tabulador * 1) & "Campo Vacio. OPTATIVO..CONTINUA PROCESAMIENTO"
                    Else
                        Flog.writeline Espacios(Tabulador * 1) & "Cotizacion Acc de Trabajo Obtenida"
                    End If
                    Flog.writeline
                    
                    'Bonificacion Ley 15.386
                    Flog.writeline "Procesando Campo 48: Bonificacion Ley 15.386"
                    
                    If arreglo(48) < 0 Then
                        Flog.writeline Espacios(Tabulador * 1) & "ERROR Bonificacion Ley 15.386 no puede ser negativa"
                        HuboError = True
                    End If
                    
                    If arreglo(48) = 0 Then
                        Flog.writeline Espacios(Tabulador * 1) & "Campo Vacio. OPTATIVO..CONTINUA PROCESAMIENTO"
                    Else
                        Flog.writeline Espacios(Tabulador * 1) & "Cotizacion Acc de Trabajo Obtenida"
                    End If
                    Flog.writeline
                    
                    'Descuentos por Cargas Familiares
                    Flog.writeline "Procesando Campo 49: Descuentos por Cargas Familiares"
                    
                    If arreglo(49) < 0 Then
                        Flog.writeline Espacios(Tabulador * 1) & "ERROR Descuentos por Cargas Familiares no puede ser negativo"
                        HuboError = True
                    End If
                    
                    If arreglo(49) = 0 Then
                        Flog.writeline Espacios(Tabulador * 1) & "Campo Vacio. OPTATIVO..CONTINUA PROCESAMIENTO"
                    Else
                        Flog.writeline Espacios(Tabulador * 1) & "Descuentos por Cargas Familiares Obtenido"
                    End If
                    Flog.writeline
                    
                    'Total INP
                    Flog.writeline "Procesando Campo 50: Total INP"
                        'PUEDE SER NEGATIVO . No hay restricciones lo dejo como viene
                    Flog.writeline Espacios(Tabulador * 1) & "Total INP obtenido"
                    Flog.writeline
                    
                    
                    'Codigo ISAPRE
                    Flog.writeline "Procesando Campo 51: Codigo ISAPRE"
                    
                    If arregloEstruc(51) = "0" Then
                        Flog.writeline Espacios(Tabulador * 1) & "ERROR No se encontro el codigo ISAPRE"
                        HuboError = True
                    Else
                    
                        If IsNumeric(arregloEstruc(51)) = True Then
                            Flog.writeline Espacios(Tabulador * 1) & "Codigo ISAPRE Obtenido "
                        Else
                            Flog.writeline Espacios(Tabulador * 1) & "ERROR. El campo contiene valores alfanuméricos y se requieren valores numéricos"
                            arregloEstruc(51) = "0"
                            HuboError = True
                        End If
                        
                    End If
                    
                    Flog.writeline
                    
                    'FONASA No se trata como una ISAPRE HJI 25/06/08
                    If arregloEstruc(51) = "07" Then
                        EsFonasa = True
                        Flog.writeline "Procesando Campo 52: Fonasa No Informa Moneda del Plan Pactado, Informando Codigo Sin ISAPRE"
                        arreglo(52) = 0
                    Else
                        'Moneda del Plan Pactado
                        Flog.writeline "Procesando Campo 52: Moneda del Plan Pactado"
                    
                        If arreglo(52) = -1 Then
                            Flog.writeline Espacios(Tabulador * 1) & "ERROR No se encontro la Moneda del Plan Pactado"
                            HuboError = True
                        Else
                            If IsNumeric(arreglo(52)) = True Then
                                Flog.writeline Espacios(Tabulador * 1) & "Moneda del Plan Pactado Obtenido "
                            Else
                                Flog.writeline Espacios(Tabulador * 1) & "ERROR. El campo contiene valores alfanuméricos y se requieren valores numéricos"
                                arregloEstruc(52) = "0"
                                HuboError = True
                            End If
                            
                        End If
                    End If
                    
                    Flog.writeline
                    
                    'FONASA No se trata como una ISAPRE HJI 25/06/08
                    If EsFonasa = True Then
                        Flog.writeline "Procesando Campo 53: Fonasa No Informa Cotizacion Obligatoria ISAPRE"
                        arreglo(53) = 0
                    Else
                        'Cotizacion Obligatoria ISAPRE
                        Flog.writeline "Procesando Campo 53: Cotizacion Obligatoria ISAPRE"
                        
                        If arreglo(53) < 0 Then
                            Flog.writeline Espacios(Tabulador * 1) & "ERROR la Cotizacion Obligatoria no puede ser negativa"
                            HuboError = True
                        End If
                        
                        If arreglo(53) = 0 Then
                            Flog.writeline Espacios(Tabulador * 1) & "ERROR. Campo Vacio. OBLIGATORIO"
                            HuboError = True
                        Else
                            Flog.writeline Espacios(Tabulador * 1) & "Cotizacion Obligatoria Isapre Obtenida"
                        End If
                    End If
                        
                    Flog.writeline
                    
                    'Bonificacion ley 58.566        CAMPO DESHABILITADO
                    Flog.writeline "Procesando Campo 54: Bonificacion ley 58.566"
                    Flog.writeline Espacios(Tabulador * 1) & "Campo DESHABILITADO "
                    Flog.writeline
                    
                    'Cotizacion Adicional Voluntaria
                    Flog.writeline "Procesando Campo 55: Cotizacion Adicional Voluntaria"
                    
                    If arreglo(55) < 0 Then
                        Flog.writeline Espacios(Tabulador * 1) & "ERROR la Cotizacion Adicional Voluntaria no puede ser negativa"
                        HuboError = True
                    End If
                    
                    If arreglo(55) = 0 Then
                        Flog.writeline Espacios(Tabulador * 1) & "Campo Vacio. OPTATIVO..CONTINUA PROCESAMIENTO"
                    Else
                        Flog.writeline Espacios(Tabulador * 1) & "Cotizacion Adicional Voluntaria Obtenida"
                    End If
                    Flog.writeline
                    
                    
                    'Otros Aportes Isapre
                    Flog.writeline "Procesando Campo 56: Otros Aportes Isapre"
                    
                    If arreglo(56) < 0 Then
                        Flog.writeline Espacios(Tabulador * 1) & "ERROR el campo no puede ser negativo"
                        HuboError = True
                    End If
                    
                    If arreglo(56) = 0 Then
                        Flog.writeline Espacios(Tabulador * 1) & "Campo Vacio. OPTATIVO..CONTINUA PROCESAMIENTO"
                    Else
                        Flog.writeline Espacios(Tabulador * 1) & "Otros Aportes Isapre Obtenidos"
                    End If
                    Flog.writeline
                    
                    
                    'Cotizacion Pactada
                    Flog.writeline "Procesando Campo 57: Cotizacion Pactada"
                    
                    If arreglo(57) < 0 Then
                        Flog.writeline Espacios(Tabulador * 1) & "ERROR el campo no puede ser negativo"
                        HuboError = True
                    End If
                    
                    If arreglo(57) = 0 Then
                        Flog.writeline Espacios(Tabulador * 1) & "Campo Vacio. OPTATIVO..CONTINUA PROCESAMIENTO"
                    Else
                        Flog.writeline Espacios(Tabulador * 1) & "Cotizacion Pactada Obtenida"
                    End If
                    Flog.writeline
                    
                    
                    'Total a Pagar Isapre
                    Flog.writeline "Procesando Campo 58: Total a Pagar Isapre"
                    
                    If arreglo(58) < 0 Then
                        Flog.writeline Espacios(Tabulador * 1) & "ERROR el campo no puede ser negativo"
                        HuboError = True
                    End If
                    
                    If arreglo(58) = 0 Then
                        Flog.writeline Espacios(Tabulador * 1) & "Campo Vacío OBLIGATORIO. Se completa con 0."
                    Else
                        Flog.writeline Espacios(Tabulador * 1) & "Total a Pagar Isapre Obtenido"
                    End If
                    Flog.writeline
                    
                    
                    'FUN DEL EMPLEADO
                    Flog.writeline "Procesando Campo 59: FUN  "
                    StrSql = " SELECT nrodoc FROM tercero " & _
                             " INNER JOIN ter_doc  ON (tercero.ternro = ter_doc.ternro) " & _
                             " INNER JOIN tipodocu on tipodocu.tidnro = ter_doc.tidnro and tipodocu.tidsigla='Fun' " & _
                             " WHERE tercero.ternro= " & rs_Empleados!ternro
                    OpenRecordset StrSql, rs_Rut
                    
                    If Not rs_Rut.EOF Then
                        FUN = rs_Rut!nrodoc
                        Flog.writeline Espacios(Tabulador * 1) & "FUN Obtenido"
                    Else
                        Flog.writeline Espacios(Tabulador * 1) & "No se encontro Numero de FUN"
                        FUN = 0
                    End If
                    Flog.writeline
                    
                    'Codigo CCAF
                    Flog.writeline "Procesando Campo 60: Codigo CCAF"
                    
                    If arregloEstruc(60) = "0" Then
                        Flog.writeline Espacios(Tabulador * 1) & "Campo Vacio. OPTATIVO..CONTINUA PROCESAMIENTO"
                    Else
                        If IsNumeric(arregloEstruc(60)) = True Then
                            Flog.writeline Espacios(Tabulador * 1) & "Codigo CCAF Obtenido"
                        Else
                            Flog.writeline Espacios(Tabulador * 1) & "ERROR. El campo contiene valores alfanuméricos y se requieren valores numéricos"
                            arregloEstruc(60) = "0"
                            HuboError = True
                        End If
                    End If
                    Flog.writeline
                    
                    'Creditos Personales
                    Flog.writeline "Procesando Campo 61: Creditos Personales"
                    
                    If arreglo(61) < 0 Then
                        Flog.writeline Espacios(Tabulador * 1) & "ERROR el campo no puede ser negativo"
                        HuboError = True
                    End If
                    
                    If arreglo(61) = 0 Then
                        Flog.writeline Espacios(Tabulador * 1) & "Campo Vacio. OPTATIVO..CONTINUA PROCESAMIENTO"
                    Else
                        Flog.writeline Espacios(Tabulador * 1) & "Creditos Personales Obtenido"
                    End If
                    Flog.writeline
                    
                    'Convenio Dental
                    Flog.writeline "Procesando Campo 62: Convenio Dental"
                    
                    If arreglo(62) < 0 Then
                        Flog.writeline Espacios(Tabulador * 1) & "ERROR el campo no puede ser negativo"
                        HuboError = True
                    End If
                    
                    If arreglo(62) = 0 Then
                        Flog.writeline Espacios(Tabulador * 1) & "Campo Vacio. OPTATIVO..CONTINUA PROCESAMIENTO"
                    Else
                        Flog.writeline Espacios(Tabulador * 1) & "Convenio Dental Obtenido"
                    End If
                    Flog.writeline
                    
                    'Descuento Por Leasing
                    Flog.writeline "Procesando Campo 63: Descuento Por Leasing"
                    
                    If arreglo(63) < 0 Then
                        Flog.writeline Espacios(Tabulador * 1) & "ERROR el campo no puede ser negativo"
                        HuboError = True
                    End If
                    
                    If arreglo(63) = 0 Then
                        Flog.writeline Espacios(Tabulador * 1) & "Campo Vacio. OPTATIVO..CONTINUA PROCESAMIENTO"
                    Else
                        Flog.writeline Espacios(Tabulador * 1) & "Descuento Por Leasing Obtenido"
                    End If
                    Flog.writeline
                                    
                    'Seguro de Vida
                    Flog.writeline "Procesando Campo 64: Seguro de Vida"
                    
                    If arreglo(64) < 0 Then
                        Flog.writeline Espacios(Tabulador * 1) & "ERROR el campo no puede ser negativo"
                        HuboError = True
                    End If
                    
                    If arreglo(64) = 0 Then
                        Flog.writeline Espacios(Tabulador * 1) & "Campo Vacio. OPTATIVO..CONTINUA PROCESAMIENTO"
                    Else
                        Flog.writeline Espacios(Tabulador * 1) & "Seguro de Vida Obtenido"
                    End If
                    Flog.writeline
                    
                    
                    'Otros CCAF
                    Flog.writeline "Procesando Campo 65: Otros CCAF"
                    
                    If arreglo(65) < 0 Then
                        Flog.writeline Espacios(Tabulador * 1) & "ERROR el campo no puede ser negativo"
                        HuboError = True
                    End If
                    
                    If arreglo(65) = 0 Then
                        Flog.writeline Espacios(Tabulador * 1) & "Campo Vacio. OPTATIVO..CONTINUA PROCESAMIENTO"
                    Else
                        Flog.writeline Espacios(Tabulador * 1) & "Otros CCAF Obtenido"
                    End If
                    Flog.writeline
                    
                    
                    'Cotizacion no Afiliado a ISAPRE
                    Flog.writeline "Procesando Campo 66: Cotizacion no Afiliado a ISAPRE"
                    
                    If arreglo(66) < 0 Then
                        Flog.writeline Espacios(Tabulador * 1) & "ERROR el campo no puede ser negativo"
                        HuboError = True
                    End If
                    
                    If arreglo(66) = 0 Then
                        Flog.writeline Espacios(Tabulador * 1) & "Campo Vacio. OPTATIVO..CONTINUA PROCESAMIENTO"
                    Else
                        Flog.writeline Espacios(Tabulador * 1) & "Cotizacion no Afiliado a ISAPRE Obtenida"
                    End If
                    Flog.writeline
                    
                    'Descuento Por cargas Familiares
                    Flog.writeline "Procesando Campo 67: Descuento Por cargas Familiares"
                    
                    If arreglo(67) = 0 Then
                        Flog.writeline Espacios(Tabulador * 1) & "Campo Vacio. OPTATIVO..CONTINUA PROCESAMIENTO"
                    Else
                        Flog.writeline Espacios(Tabulador * 1) & "Descuento Por cargas Familiares Obtenido"
                    End If
                    Flog.writeline
                    
                    'Codigo Mutual
                    Flog.writeline "Procesando Campo 68: Codigo Mutual"
                    
                    If arregloEstruc(68) = "0" Then
                        Flog.writeline Espacios(Tabulador * 1) & "Campo Vacio. OPTATIVO..CONTINUA PROCESAMIENTO"
                    Else
                        If IsNumeric(arregloEstruc(68)) = True Then
                            Flog.writeline Espacios(Tabulador * 1) & "Codigo Mutual Obtenido"
                        Else
                            Flog.writeline Espacios(Tabulador * 1) & "ERROR. El campo contiene valores alfanuméricos y se requieren valores numéricos"
                            arregloEstruc(68) = "0"
                            HuboError = True
                        End If
                    End If
                    Flog.writeline
                    
                    'Cotizacion ACC de Trabajo
                    Flog.writeline "Procesando Campo 69: Cotizacion ACC de Trabajo"
                    
                    If arreglo(69) < 0 Then
                        Flog.writeline Espacios(Tabulador * 1) & "ERROR el campo no puede ser negativo"
                        HuboError = True
                    End If
                    
                    If arreglo(69) = 0 Then
                        Flog.writeline Espacios(Tabulador * 1) & "Campo Vacio. OPTATIVO..CONTINUA PROCESAMIENTO"
                    Else
                        Flog.writeline Espacios(Tabulador * 1) & "Cotizacion ACC de Trabajo Obtenido"
                    End If
                    Flog.writeline
                    
                   'Sucursal para pago mutual
                    Flog.writeline "Procesando Campo 70: Sucursal para pago mutual"
                    
                    If arregloEstruc(70) = "0" Then
                        Flog.writeline Espacios(Tabulador * 1) & "Campo Vacio. OPTATIVO..CONTINUA PROCESAMIENTO"
                    Else
                        If IsNumeric(arregloEstruc(70)) = True Then
                            Flog.writeline Espacios(Tabulador * 1) & "Sucursal para pago mutual Obtenida"
                        Else
                            Flog.writeline Espacios(Tabulador * 1) & "ERROR. El campo contiene valores alfanuméricos y se requieren valores numéricos"
                            arregloEstruc(70) = "0"
                            HuboError = True
                        End If
                        
                    End If
                    Flog.writeline
                  
                   'Institucion Autorizada Ahorro Previsional
                    Flog.writeline "Procesando Campo 71: Institucion Autorizada Ahorro Previsional"
                    
                    If arregloEstruc(71) = "0" Then
                        Flog.writeline Espacios(Tabulador * 1) & "ERROR. Campo Vacio. OBLIGATORIO"
                        HuboError = True
                    Else
                        If IsNumeric(arregloEstruc(71)) = True Then
                            Flog.writeline Espacios(Tabulador * 1) & "Institucion Autorizada Ahorro Previsional Obtenido"
                        Else
                            Flog.writeline Espacios(Tabulador * 1) & "ERROR. El campo contiene valores alfanuméricos y se requieren valores numéricos"
                            arregloEstruc(71) = "0"
                            HuboError = True
                        End If
                    End If
                    Flog.writeline
                   
                   'Forma de PAGO APV
                   Flog.writeline "Procesando Campo 72: Forma de pago Ahorro Previsional Voluntario"
                   If arregloEstruc(72) = "0" And arregloEstruc(71) <> "000" Then
                        Flog.writeline Espacios(Tabulador * 1) & "ERROR. Campo Vacio. OBLIGATORIO"
                        HuboError = True
                    Else
                        If IsNumeric(arregloEstruc(72)) = True Then
                            Flog.writeline Espacios(Tabulador * 1) & "Forma de pago Ahorro Previsional Voluntario Obtenido"
                        Else
                            Flog.writeline Espacios(Tabulador * 1) & "ERROR. El campo contiene valores alfanuméricos y se requieren valores numéricos"
                            arregloEstruc(72) = "0"
                            HuboError = True
                        End If
                        
                    End If
                    Flog.writeline
                   
                    'Cotizacion Ahorro Previsional Voluntario
                    Flog.writeline "Procesando Campo 73: Cotizacion Ahorro Previsional Voluntario"
                    
                    If arreglo(73) < 0 Then
                        Flog.writeline Espacios(Tabulador * 1) & "ERROR el campo no puede ser negativo"
                        HuboError = True
                    End If
                    
                    If arreglo(73) = 0 Then
                        Flog.writeline Espacios(Tabulador * 1) & "Campo Vacio. OPTATIVO..CONTINUA PROCESAMIENTO"
                    Else
                        Flog.writeline Espacios(Tabulador * 1) & "Cotizacion Ahorro Previsional Voluntario Obtenido"
                    End If
                    Flog.writeline
                   
                   'Cotizacion Depositos Convenidos
                    Flog.writeline "Procesando Campo 74: Cotizacion Depositos Convenidos"
                    
                    If arreglo(74) < 0 Then
                        Flog.writeline Espacios(Tabulador * 1) & "ERROR el campo no puede ser negativo"
                        HuboError = True
                    End If
                    
                    If arreglo(74) = 0 Then
                        Flog.writeline Espacios(Tabulador * 1) & "Campo Vacio. OPTATIVO..CONTINUA PROCESAMIENTO"
                    Else
                        Flog.writeline Espacios(Tabulador * 1) & "Cotizacion Depositos Convenidos Obtenido"
                    End If
                    Flog.writeline
                   
                   'Renta imponible Seguro Cesantia
                    Flog.writeline "Procesando Campo 75: Renta imponible Seguro Cesantia"
                    
                    If arreglo(75) < 0 Then
                        Flog.writeline Espacios(Tabulador * 1) & "ERROR el campo no puede ser negativo"
                        HuboError = True
                    End If
                    
                    SeguroCesantia = False
                    If arreglo(75) = 0 Then
                        Flog.writeline Espacios(Tabulador * 1) & "Campo Vacio. OPTATIVO..CONTINUA PROCESAMIENTO"
                    Else
                        Flog.writeline Espacios(Tabulador * 1) & "Renta imponible Seguro Cesantia Obtenido"
                        SeguroCesantia = True
                    End If
                    Flog.writeline
                    
                    'Aporte Trabajador Seguro Cesantia
                    Flog.writeline "Procesando Campo 76: Aporte Trabajador Seguro Cesantia"
                    
                    If arreglo(76) < 0 Then
                        Flog.writeline Espacios(Tabulador * 1) & "ERROR el campo no puede ser negativo"
                        HuboError = True
                    End If
                    
                    If arreglo(76) = 0 Then
                        Flog.writeline Espacios(Tabulador * 1) & "Campo Vacio. OPTATIVO..CONTINUA PROCESAMIENTO"
                    Else
                        Flog.writeline Espacios(Tabulador * 1) & "Aporte Trabajador Seguro Cesantia Obtenido"
                    End If
                    Flog.writeline
                    
                    'Aporte Empleador Seguro Cesantia
                    Flog.writeline "Procesando Campo 77: Aporte Empleador Seguro Cesantia"
                    
                    If arreglo(77) < 0 Then
                        Flog.writeline Espacios(Tabulador * 1) & "ERROR el campo no puede ser negativo"
                        HuboError = True
                    End If
                    
                    If arreglo(77) = 0 Then
                        Flog.writeline Espacios(Tabulador * 1) & "Campo Vacio. OPTATIVO..CONTINUA PROCESAMIENTO"
                    Else
                        Flog.writeline Espacios(Tabulador * 1) & "Aporte Empleador Seguro Cesantia Obtenido"
                    End If
                    Flog.writeline
                    
                    'Centro de Costo
                    Flog.writeline "Procesando Campo 78: Centro de Costo"
                    ' 08-03-07 Diego Rosso
                    If arregloEstruc(78) = "0" Then
                        Flog.writeline Espacios(Tabulador * 1) & "Campo Vacio. OPTATIVO..CONTINUA PROCESAMIENTO"
                        
                    Else
                        Flog.writeline Espacios(Tabulador * 1) & "Centro de Costo Obtenido"
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
                    StrSql = "INSERT INTO rep_previred (bpronro, ternro, num_linea, Titulo, pliqnro_Desde, pliqnro_hasta, empnro, rut, DV, Apellido, Apellido2, Nombres, sexo, tipo_pago, Periodo_desde,"
                    StrSql = StrSql & "Periodo_hasta, renta_imp, reg_pre, TipTrabajador, DiasTrab, CodmovPer, fechadesde, fechahasta, TramoAsigFam, NumCargasSim, NumCargasMat, "
                    StrSql = StrSql & "NumCargasInv, AsigFam, AsigFamRetro, ReintCarFam, CodAFP, CotizObligAFP, CAVoluntaAFP, RenImpSustAFP, TasaPact, AportIndem, NumPeriodos, "
                    StrSql = StrSql & "PeriDesdeAFP, PeriHastaAFP, PuesTrabPesado, PorcCotizTrabPesa, CotizTrabPesa, RutPag, DVPag, CodCaReg, TasaCotCajPrev, CotizObligINP, "
                    StrSql = StrSql & "RentImpoDesah, CodCaRegDesah, TasaCotDesah, CotizDesah, OtroINP, CotizFonasa, CotizAccTrab, BonLeyInp, DescCargFam, TotPagINP, "
                    StrSql = StrSql & "CodInstSal, MonPlanIsapre, CotizObligIsapre, BonifLeyIsapre, CotizAdicVolun, OtrosIsapre, CotizPact, TotPagIsapre, NumFun, CodCCAF, CredPerCCAF,"
                    StrSql = StrSql & "DescDentCCAF, DescLeasCCAF, DescVidaCCAF, OtrosDesCCAF, CotCCAFnoIsapre, DesCarFamCCAF, CodMut, CotizAccTrabMut, SucPagMut "
                    StrSql = StrSql & ",InstAutAPV, ForPagAPV, CotizAPV, CotizDepConv, RentTotImp, AporTrabSeg, AporEmpSeg, CentroCosto, auxdeci"
                    
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
                    StrSql = StrSql & TipoPago & ","
                    StrSql = StrSql & "'" & Month(Fechadesde) & Year(Fechadesde) & "',"
                    StrSql = StrSql & "'" & Month(Fechahasta) & Year(Fechahasta) & "',"
                    StrSql = StrSql & arreglo(10) & "," 'Renta imponible
                    StrSql = StrSql & "'" & Mid(arregloEstruc(11), 1, 50) & "'," 'Regimen Previsional
                    StrSql = StrSql & arregloEstruc(12) & "," 'Tipo Trabajador
                    StrSql = StrSql & arreglo(13) & "," 'Dias Trabajados
                    StrSql = StrSql & arregloMov(1) & "," 'Codigo Movimiento de personal
                    StrSql = StrSql & IIf(arregloFecD(1) = vbNull, "Null", ConvFecha(arregloFecD(1))) & "," 'Fecha Desde para el movimiento
                    StrSql = StrSql & IIf(arregloFecH(1) = vbNull, "Null", ConvFecha(arregloFecH(1))) & "," 'Fecha Hasta para el movimiento
                    StrSql = StrSql & "'" & arregloEstruc(17) & "'," 'Tramo Asignacion Familiar
                    StrSql = StrSql & arreglo(18) & "," 'Numero de cargas simples
                    StrSql = StrSql & arreglo(19) & "," 'Numero de cargas Maternales
                    StrSql = StrSql & arreglo(20) & "," 'Numero de cargas Invalidas
                    StrSql = StrSql & arreglo(21) & "," 'ASignacion Familiar
                    StrSql = StrSql & arreglo(22) & "," 'ASignacion Familiar Retroactiva
                    StrSql = StrSql & arreglo(23) & "," 'Renta Carga Familiares
                    StrSql = StrSql & arregloEstruc(24) & "," 'Codigo AFP
                    StrSql = StrSql & arreglo(25) & "," 'Cotizacion Obligatoria AFP
                    StrSql = StrSql & arreglo(26) & "," 'Cuenta de ahorro Voluntario
                    StrSql = StrSql & arreglo(27) & "," 'Renta Imponible sust AFP
                    StrSql = StrSql & arreglo(28) & "," 'Tasa Pactada
                    StrSql = StrSql & arreglo(29) & "," 'Aporte Indem
                    StrSql = StrSql & 0 & "," 'Numero de Periodos REVISAR!!!!!!!!!!!!!!!!!!!!!
                    StrSql = StrSql & "Null" & ","  'REVISAR CAMBIAR
                    StrSql = StrSql & "Null" & ","  'REVISAR CAMBIAR
                    StrSql = StrSql & "'" & Mid(arregloEstruc(33), 1, 40) & "'," 'Puesto de trabajo Pesado
                    StrSql = StrSql & arreglo(34) & "," 'Porcentaje Cotizacion Trabajo Pesado
                    StrSql = StrSql & arreglo(35) & "," 'Cotizacion Trabajo Pesado
                    StrSql = StrSql & "Null" & "," 'RUT Empresa REVISAR!!!!!!!!!!!!!!!!!!!!!
                    StrSql = StrSql & "Null" & "," 'DV Empresa REVISAR!!!!!!!!!!!!!!!!!!!!!
                    StrSql = StrSql & arregloEstruc(38) & "," 'Codigo Caja Regimen
                    StrSql = StrSql & arreglo(39) & "," 'Tasa cotizacion EX cajas de Regimen
                    StrSql = StrSql & arreglo(40) & "," 'Cotizacion Obligatoria INP
                    StrSql = StrSql & arreglo(41) & "," 'Renta Imponible Desahucio
                    StrSql = StrSql & arregloEstruc(42) & "," 'COdigo ex caja Regimen
                    StrSql = StrSql & arreglo(43) & "," 'Tasa Cotizacion Desahucio
                    StrSql = StrSql & arreglo(44) & "," 'Cotizacion Desahucio
                    StrSql = StrSql & 0 & ","    'Otros Aportes INP
                    StrSql = StrSql & arreglo(46) & "," 'Cotizacion Fonasa
                    StrSql = StrSql & arreglo(47) & "," 'Cotizacion Accidente de Trabajo
                    StrSql = StrSql & arreglo(48) & "," 'Bonificacion ley 15.386
                    StrSql = StrSql & arreglo(49) & "," 'Descuento por Cargas Familiares
                    StrSql = StrSql & arreglo(50) & "," 'Total a Pagar al INP
                    StrSql = StrSql & arregloEstruc(51) & "," 'Codigo Institucion de Salud
                    StrSql = StrSql & arreglo(52) & "," 'Moneda del plan pactado con Isapre
                    StrSql = StrSql & arreglo(53) & "," 'Cotizacion Obligatoria ISAPRE
                    StrSql = StrSql & arreglo(54) & "," 'Bonificacion ley 18.566
                    StrSql = StrSql & arreglo(55) & "," 'Cotizacion adicional Voluntaria
                    StrSql = StrSql & arreglo(56) & "," 'Otros ISAPRE
                    StrSql = StrSql & arreglo(57) & "," 'Cotizacion Pactada
                    StrSql = StrSql & arreglo(58) & "," 'Total a pagar Isapre
                    StrSql = StrSql & FUN & ","  'FUN
                    StrSql = StrSql & arregloEstruc(60) & "," 'Codigo CCAF
                    StrSql = StrSql & arreglo(61) & "," 'Creditos Personales CCAF
                    StrSql = StrSql & arreglo(62) & "," 'Descuento Dental
                    StrSql = StrSql & arreglo(63) & "," 'Descuento por Leasing
                    StrSql = StrSql & arreglo(64) & "," 'Descuentos por Seguro de Vida
                    StrSql = StrSql & arreglo(65) & "," 'Otros Descuentos CCAF
                    StrSql = StrSql & arreglo(66) & "," 'Cotiz CCAF de no Afiliados ISAPRE
                    StrSql = StrSql & arreglo(67) & "," 'Descuentos Cargas Familiares CCAF
                    StrSql = StrSql & arregloEstruc(68) & "," 'Codigo Mutual
                    StrSql = StrSql & arreglo(69) & "," 'Cotizacion Accidente del trabajo
                    StrSql = StrSql & arregloEstruc(70) & "," 'Sucursal Para Pago Mutual
                    StrSql = StrSql & arregloEstruc(71) & "," 'Inst Autor Ahorro Previsonal Vol
                    StrSql = StrSql & arregloEstruc(72) & ","  'Forma de Pago APV
                    StrSql = StrSql & arreglo(73) & "," 'Cotizacion Ahorro Previsonal Vol
                    StrSql = StrSql & arreglo(74) & "," 'Cotizacion Depositos convenidos
                    StrSql = StrSql & IIf(EstaINP And SeguroCesantia, 0, arreglo(75)) & "," 'Renta total Imponible Seguro de Cesantia
                    StrSql = StrSql & IIf(EstaINP And SeguroCesantia, 0, arreglo(76)) & "," 'Aporte Trabajador Seguro de Cesantia
                    StrSql = StrSql & IIf(EstaINP And SeguroCesantia, 0, arreglo(77)) & "," 'Aporte Empleador Seguro de Cesantia
                    StrSql = StrSql & "'" & Mid(arregloEstruc(78), 1, 3) & "',"   'Centro Costo, sucursal, etc
                    StrSql = StrSql & IIf(HuboError, -1, 0)
                                        
                    StrSql = StrSql & ")"
                    Flog.writeline
                    Flog.writeline "Insertando : " & StrSql
                    objConn.Execute StrSql, , adExecuteNoRecords
                    Flog.writeline
                    Flog.writeline
                    'Sumo el numero de linea
                    Num_linea = Num_linea + 1
                    CantEmplSinError = CantEmplSinError + 1
                    Flog.writeline
                    Flog.writeline "SE GRABO EL EMPLEADO "
                    Flog.writeline
                    
                    aux = 2
                    Do While aux <= total_mov

                   'Inserto en rep_previred las lineas adicionales
                    StrSql = "INSERT INTO rep_previred (bpronro, ternro, num_linea, Titulo, pliqnro_Desde, pliqnro_hasta, empnro, rut, DV, Apellido, Apellido2, Nombres, sexo, tipo_pago, Periodo_desde,"
                    StrSql = StrSql & "Periodo_hasta, renta_imp, reg_pre, TipTrabajador, DiasTrab, CodmovPer, fechadesde, fechahasta, TramoAsigFam, NumCargasSim, NumCargasMat, "
                    StrSql = StrSql & "NumCargasInv, AsigFam, AsigFamRetro, ReintCarFam, CodAFP, CotizObligAFP, CAVoluntaAFP, RenImpSustAFP, TasaPact, AportIndem, NumPeriodos, "
                    StrSql = StrSql & "PeriDesdeAFP, PeriHastaAFP, PuesTrabPesado, PorcCotizTrabPesa, CotizTrabPesa, RutPag, DVPag, CodCaReg, TasaCotCajPrev, CotizObligINP, "
                    StrSql = StrSql & "RentImpoDesah, CodCaRegDesah, TasaCotDesah, CotizDesah, OtroINP, CotizFonasa, CotizAccTrab, BonLeyInp, DescCargFam, TotPagINP, "
                    StrSql = StrSql & "CodInstSal, MonPlanIsapre, CotizObligIsapre, BonifLeyIsapre, CotizAdicVolun, OtrosIsapre, CotizPact, TotPagIsapre, NumFun, CodCCAF, CredPerCCAF,"
                    StrSql = StrSql & "DescDentCCAF, DescLeasCCAF, DescVidaCCAF, OtrosDesCCAF, CotCCAFnoIsapre, DesCarFamCCAF, CodMut, CotizAccTrabMut, SucPagMut "
                    StrSql = StrSql & ",InstAutAPV, ForPagAPV, CotizAPV, CotizDepConv, RentTotImp, AporTrabSeg, AporEmpSeg, CentroCosto, auxdeci"
                    
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
                    StrSql = StrSql & TipoPago & ","
                    StrSql = StrSql & "'" & Month(Fechadesde) & Year(Fechadesde) & "',"
                    StrSql = StrSql & "'" & Month(Fechahasta) & Year(Fechahasta) & "',"
                    StrSql = StrSql & "0," 'Renta imponible
                    StrSql = StrSql & "'" & Mid(arregloEstruc(11), 1, 50) & "'," 'Regimen Previsional
                    StrSql = StrSql & arregloEstruc(12) & "," 'Tipo Trabajador
                    StrSql = StrSql & "0," 'Dias Trabajados
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
                    StrSql = StrSql & arregloEstruc(24) & "," 'Codigo AFP
                    StrSql = StrSql & "0," 'Cotizacion Obligatoria AFP
                    StrSql = StrSql & "0," 'Cuenta de ahorro Voluntario
                    StrSql = StrSql & "0," 'Renta Imponible sust AFP
                    StrSql = StrSql & "0.00," 'Tasa Pactada
                    StrSql = StrSql & "0," 'Aporte Indem
                    StrSql = StrSql & "00," 'Numero de Periodos REVISAR!!!!!!!!!!!!!!!!!!!!!
                    StrSql = StrSql & "Null" & ","  'REVISAR CAMBIAR
                    StrSql = StrSql & "Null" & ","  'REVISAR CAMBIAR
                    StrSql = StrSql & "Null" & ","  'Puesto de trabajo Pesado
                    StrSql = StrSql & "00.00," 'Porcentaje Cotizacion Trabajo Pesado
                    StrSql = StrSql & "0," 'Cotizacion Trabajo Pesado
                    StrSql = StrSql & "Null" & "," 'RUT Empresa REVISAR!!!!!!!!!!!!!!!!!!!!!
                    StrSql = StrSql & "Null" & "," 'DV Empresa REVISAR!!!!!!!!!!!!!!!!!!!!!
                    StrSql = StrSql & arregloEstruc(38) & "," 'Codigo Caja Regimen
                    StrSql = StrSql & "00.00," 'Tasa cotizacion EX cajas de Regimen
                    StrSql = StrSql & "0," 'Cotizacion Obligatoria INP
                    StrSql = StrSql & "0," 'Renta Imponible Desahucio
                    StrSql = StrSql & "0," 'COdigo ex caja Regimen
                    StrSql = StrSql & "0," 'Tasa Cotizacion Desahucio
                    StrSql = StrSql & "0," 'Cotizacion Desahucio
                    StrSql = StrSql & "0,"    'Otros Aportes INP
                    StrSql = StrSql & "0," 'Cotizacion Fonasa
                    StrSql = StrSql & "0," 'Cotizacion Accidente de Trabajo
                    StrSql = StrSql & "0," 'Bonificacion ley 15.386
                    StrSql = StrSql & "0," 'Descuento por Cargas Familiares
                    StrSql = StrSql & "0," 'Total a Pagar al INP
                    StrSql = StrSql & arregloEstruc(51) & "," 'Codigo Institucion de Salud
                    StrSql = StrSql & "0," 'Moneda del plan pactado con Isapre
                    StrSql = StrSql & "0," 'Cotizacion Obligatoria ISAPRE
                    StrSql = StrSql & "0," 'Bonificacion ley 18.566
                    StrSql = StrSql & "0," 'Cotizacion adicional Voluntaria
                    StrSql = StrSql & "0," 'Otros ISAPRE
                    StrSql = StrSql & "0," 'Cotizacion Pactada
                    StrSql = StrSql & "0," 'Total a pagar Isapre
                    StrSql = StrSql & "0,"  'FUN
                    StrSql = StrSql & arregloEstruc(60) & "," 'Codigo CCAF
                    StrSql = StrSql & "0," 'Creditos Personales CCAF
                    StrSql = StrSql & "0," 'Descuento Dental
                    StrSql = StrSql & "0," 'Descuento por Leasing
                    StrSql = StrSql & "0," 'Descuentos por Seguro de Vida
                    StrSql = StrSql & "0," 'Otros Descuentos CCAF
                    StrSql = StrSql & "0," 'Cotiz CCAF de no Afiliados ISAPRE
                    StrSql = StrSql & "0," 'Descuentos Cargas Familiares CCAF
                    StrSql = StrSql & arregloEstruc(68) & "," 'Codigo Mutual
                    StrSql = StrSql & "0," 'Cotizacion Accidente del trabajo
                    StrSql = StrSql & arregloEstruc(70) & "," 'Sucursal Para Pago Mutual
                    StrSql = StrSql & "000," 'Inst Autor Ahorro Previsonal Vol
                    StrSql = StrSql & "0,"  'Forma de Pago APV
                    StrSql = StrSql & "0," 'Cotizacion Ahorro Previsonal Vol
                    StrSql = StrSql & "0," 'Cotizacion Depositos convenidos
                    StrSql = StrSql & "0," 'Renta total Imponible Seguro de Cesantia
                    StrSql = StrSql & "0," 'Aporte Trabajador Seguro de Cesantia
                    StrSql = StrSql & "0," 'Aporte Empleador Seguro de Cesantia
                    StrSql = StrSql & "'" & Mid(arregloEstruc(78), 1, 3) & "',"   'Centro Costo, sucursal, etc
                    StrSql = StrSql & IIf(HuboError, -1, 0)
                    
                    StrSql = StrSql & ")"
                    Flog.writeline
                    Flog.writeline "Insertando : " & StrSql
                    objConn.Execute StrSql, , adExecuteNoRecords
                    Flog.writeline
                    Flog.writeline
                    'Sumo el numero de linea
                    Num_linea = Num_linea + 1
                    Flog.writeline
                    Flog.writeline "SE GRABO LINEA ADICIONAL"
                    Flog.writeline

                    aux = aux + 1
                    Loop
                    
                    Case 2: 'Gratificaciones
                    
                    
                    MesDesdeRec = ""
                    AnioDesdeRec = ""
                    MesHastaRec = ""
                    AnioHastaRec = ""
                    
                    'Busco los periodos recalculo
                    StrSql = "SELECT periodo.pliqmes, periodo.pliqanio FROM impuni_peri "
                    StrSql = StrSql & " INNER JOIN periodo ON periodo.pliqnro = impuni_peri.pliqnro"
                    StrSql = StrSql & " WHERE pronro IN (" & Lista_Pro & ")"
                    StrSql = StrSql & " ORDER BY periodo.pliqanio, periodo.pliqmes"
                    OpenRecordset StrSql, rs_Estr_cod
                    If Not rs_Estr_cod.EOF Then
                        MesDesdeRec = rs_Estr_cod!pliqmes
                        AnioDesdeRec = rs_Estr_cod!pliqanio
                    End If
                    rs_Estr_cod.Close
                    
                    StrSql = "SELECT periodo.pliqmes, periodo.pliqanio FROM impuni_peri "
                    StrSql = StrSql & " INNER JOIN periodo ON periodo.pliqnro = impuni_peri.pliqnro"
                    StrSql = StrSql & " WHERE pronro IN (" & Lista_Pro & ")"
                    StrSql = StrSql & " ORDER BY periodo.pliqanio Desc, periodo.pliqmes Desc"
                    OpenRecordset StrSql, rs_Estr_cod
                    If Not rs_Estr_cod.EOF Then
                        MesHastaRec = rs_Estr_cod!pliqmes
                        AnioHastaRec = rs_Estr_cod!pliqanio
                    End If
                    rs_Estr_cod.Close
                    
                    
                   'Inserto en rep_previred
                    StrSql = "INSERT INTO rep_previred (bpronro, ternro, num_linea, Titulo, pliqnro_Desde, pliqnro_hasta, empnro, rut, DV, Apellido, Apellido2, Nombres, sexo, tipo_pago, Periodo_desde,"
                    StrSql = StrSql & "Periodo_hasta, renta_imp, reg_pre, TipTrabajador, DiasTrab, CodmovPer, fechadesde, fechahasta, TramoAsigFam, NumCargasSim, NumCargasMat, "
                    StrSql = StrSql & "NumCargasInv, AsigFam, AsigFamRetro, ReintCarFam, CodAFP, CotizObligAFP, CAVoluntaAFP, RenImpSustAFP, TasaPact, AportIndem, NumPeriodos, "
                    StrSql = StrSql & "PeriDesdeAFP, PeriHastaAFP, PuesTrabPesado, PorcCotizTrabPesa, CotizTrabPesa, RutPag, DVPag, CodCaReg, TasaCotCajPrev, CotizObligINP, "
                    StrSql = StrSql & "RentImpoDesah, CodCaRegDesah, TasaCotDesah, CotizDesah, OtroINP, CotizFonasa, CotizAccTrab, BonLeyInp, DescCargFam, TotPagINP, "
                    StrSql = StrSql & "CodInstSal, MonPlanIsapre, CotizObligIsapre, BonifLeyIsapre, CotizAdicVolun, OtrosIsapre, CotizPact, TotPagIsapre, NumFun, CodCCAF, CredPerCCAF,"
                    StrSql = StrSql & "DescDentCCAF, DescLeasCCAF, DescVidaCCAF, OtrosDesCCAF, CotCCAFnoIsapre, DesCarFamCCAF, CodMut, CotizAccTrabMut, SucPagMut "
                    StrSql = StrSql & ",InstAutAPV, ForPagAPV, CotizAPV, CotizDepConv, RentTotImp, AporTrabSeg, AporEmpSeg, CentroCosto, auxdeci"
                    
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
                    StrSql = StrSql & TipoPago & ","
                    StrSql = StrSql & "'" & Format(MesDesdeRec, "00") & Format(AnioDesdeRec, "0000") & "',"
                    StrSql = StrSql & "'" & Format(MesHastaRec, "00") & Format(AnioHastaRec, "0000") & "',"
                    StrSql = StrSql & arreglo(10) & "," 'Renta imponible
                    StrSql = StrSql & "'" & Mid(arregloEstruc(11), 1, 50) & "'," 'Regimen Previsional
                    StrSql = StrSql & arregloEstruc(12) & "," 'Tipo Trabajador
                    StrSql = StrSql & "30," 'Dias Trabajados
                    StrSql = StrSql & "0," 'Codigo Movimiento de personal
                    StrSql = StrSql & "'          '," 'Fecha Desde para el movimiento
                    StrSql = StrSql & "'          '," 'Fecha Hasta para el movimiento
                    StrSql = StrSql & "' '," 'Tramo Asignacion Familiar
                    StrSql = StrSql & "0," 'Numero de cargas simples
                    StrSql = StrSql & "0," 'Numero de cargas Maternales
                    StrSql = StrSql & "0," 'Numero de cargas Invalidas
                    StrSql = StrSql & "0," 'ASignacion Familiar
                    StrSql = StrSql & "0," 'ASignacion Familiar Retroactiva
                    StrSql = StrSql & "0," 'Renta Carga Familiares
                    StrSql = StrSql & arregloEstruc(24) & "," 'Codigo AFP
                    StrSql = StrSql & arreglo(25) & "," 'Cotizacion Obligatoria AFP
                    StrSql = StrSql & "0," 'Cuenta de ahorro Voluntario
                    StrSql = StrSql & "0," 'Renta Imponible sust AFP
                    StrSql = StrSql & "0.00," 'Tasa Pactada
                    StrSql = StrSql & "0," 'Aporte Indem
                    StrSql = StrSql & "00," 'Numero de Periodos REVISAR!!!!!!!!!!!!!!!!!!!!!
                    StrSql = StrSql & "Null" & ","  'REVISAR CAMBIAR
                    StrSql = StrSql & "Null" & ","  'REVISAR CAMBIAR
                    StrSql = StrSql & "Null" & ","  'Puesto de trabajo Pesado
                    StrSql = StrSql & "00.00," 'Porcentaje Cotizacion Trabajo Pesado
                    StrSql = StrSql & "0," 'Cotizacion Trabajo Pesado
                    StrSql = StrSql & "0" & "," 'RUT Empresa REVISAR!!!!!!!!!!!!!!!!!!!!!
                    StrSql = StrSql & "0" & "," 'DV Empresa REVISAR!!!!!!!!!!!!!!!!!!!!!
                    StrSql = StrSql & arregloEstruc(38) & "," 'Codigo Caja Regimen
                    StrSql = StrSql & arreglo(39) & "," 'Tasa cotizacion EX cajas de Regimen
                    StrSql = StrSql & arreglo(40) & "," 'Cotizacion Obligatoria INP
                    StrSql = StrSql & arreglo(41) & "," 'Renta Imponible Desahucio
                    StrSql = StrSql & arregloEstruc(42) & "," 'COdigo ex caja Regimen
                    StrSql = StrSql & arreglo(43) & "," 'Tasa Cotizacion Desahucio
                    StrSql = StrSql & arreglo(44) & "," 'Cotizacion Desahucio
                    StrSql = StrSql & 0 & ","    'Otros Aportes INP
                    StrSql = StrSql & arreglo(46) & "," 'Cotizacion Fonasa
                    StrSql = StrSql & arreglo(47) & "," 'Cotizacion Accidente de Trabajo
                    StrSql = StrSql & "0," 'Bonificacion ley 15.386
                    StrSql = StrSql & "0," 'Descuento por Cargas Familiares
                    StrSql = StrSql & arreglo(50) & "," 'Total a Pagar al INP
                    StrSql = StrSql & arregloEstruc(51) & "," 'Codigo Institucion de Salud
                    StrSql = StrSql & "1," 'Moneda del plan pactado con Isapre
                    StrSql = StrSql & arreglo(53) & "," 'Cotizacion Obligatoria ISAPRE
                    StrSql = StrSql & "0," 'Bonificacion ley 18.566
                    StrSql = StrSql & "0," 'Cotizacion adicional Voluntaria
                    StrSql = StrSql & "0," 'Otros ISAPRE
                    StrSql = StrSql & "0," 'Cotizacion Pactada
                    StrSql = StrSql & arreglo(58) & "," 'Total a pagar Isapre
                    StrSql = StrSql & FUN & ","  'FUN
                    StrSql = StrSql & arregloEstruc(60) & "," 'Codigo CCAF
                    StrSql = StrSql & "0," 'Creditos Personales CCAF
                    StrSql = StrSql & "0," 'Descuento Dental
                    StrSql = StrSql & "0," 'Descuento por Leasing
                    StrSql = StrSql & "0," 'Descuentos por Seguro de Vida
                    StrSql = StrSql & "0," 'Otros Descuentos CCAF
                    StrSql = StrSql & arreglo(66) & "," 'Cotiz CCAF de no Afiliados ISAPRE
                    StrSql = StrSql & "0," 'Descuentos Cargas Familiares CCAF
                    StrSql = StrSql & arregloEstruc(68) & "," 'Codigo Mutual
                    StrSql = StrSql & arreglo(69) & "," 'Cotizacion Accidente del trabajo
                    StrSql = StrSql & arregloEstruc(70) & "," 'Sucursal Para Pago Mutual
                    StrSql = StrSql & "000," 'Inst Autor Ahorro Previsonal Vol
                    StrSql = StrSql & "0,"  'Forma de Pago APV
                    StrSql = StrSql & "0," 'Cotizacion Ahorro Previsonal Vol
                    StrSql = StrSql & "0," 'Cotizacion Depositos convenidos
                    StrSql = StrSql & IIf(EstaINP And SeguroCesantia, 0, arreglo(75)) & "," 'Renta total Imponible Seguro de Cesantia
                    StrSql = StrSql & IIf(EstaINP And SeguroCesantia, 0, arreglo(76)) & "," 'Aporte Trabajador Seguro de Cesantia
                    StrSql = StrSql & IIf(EstaINP And SeguroCesantia, 0, arreglo(77)) & ","
                    StrSql = StrSql & "'" & Mid(arregloEstruc(78), 1, 3) & "',"   'Centro Costo, sucursal, etc
                    StrSql = StrSql & IIf(HuboError, -1, 0)
                    
                    StrSql = StrSql & ")"
                    Flog.writeline
                    Flog.writeline "Insertando : " & StrSql
                    objConn.Execute StrSql, , adExecuteNoRecords
                    Flog.writeline
                    Flog.writeline
                    'Sumo el numero de linea
                    Num_linea = Num_linea + 1
                    Flog.writeline
                    Flog.writeline "SE GRABO LINEA DE GRATIFICACION"
                    Flog.writeline
                    
                    End Select
                    
                    If arreglo(79) <> 0 Then
                   'Hay que crear una nueva linea porque hay un nuevo APV
                    StrSql = "INSERT INTO rep_previred (bpronro, ternro, num_linea, Titulo, pliqnro_Desde, pliqnro_hasta, empnro, rut, DV, Apellido, Apellido2, Nombres, sexo, tipo_pago, Periodo_desde,"
                    StrSql = StrSql & "Periodo_hasta, renta_imp, reg_pre, TipTrabajador, DiasTrab, CodmovPer, fechadesde, fechahasta, TramoAsigFam, NumCargasSim, NumCargasMat, "
                    StrSql = StrSql & "NumCargasInv, AsigFam, AsigFamRetro, ReintCarFam, CodAFP, CotizObligAFP, CAVoluntaAFP, RenImpSustAFP, TasaPact, AportIndem, NumPeriodos, "
                    StrSql = StrSql & "PeriDesdeAFP, PeriHastaAFP, PuesTrabPesado, PorcCotizTrabPesa, CotizTrabPesa, RutPag, DVPag, CodCaReg, TasaCotCajPrev, CotizObligINP, "
                    StrSql = StrSql & "RentImpoDesah, CodCaRegDesah, TasaCotDesah, CotizDesah, OtroINP, CotizFonasa, CotizAccTrab, BonLeyInp, DescCargFam, TotPagINP, "
                    StrSql = StrSql & "CodInstSal, MonPlanIsapre, CotizObligIsapre, BonifLeyIsapre, CotizAdicVolun, OtrosIsapre, CotizPact, TotPagIsapre, NumFun, CodCCAF, CredPerCCAF,"
                    StrSql = StrSql & "DescDentCCAF, DescLeasCCAF, DescVidaCCAF, OtrosDesCCAF, CotCCAFnoIsapre, DesCarFamCCAF, CodMut, CotizAccTrabMut, SucPagMut "
                    StrSql = StrSql & ",InstAutAPV, ForPagAPV, CotizAPV, CotizDepConv, RentTotImp, AporTrabSeg, AporEmpSeg, CentroCosto, auxdeci"
                    
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
                    StrSql = StrSql & "1,"
                    StrSql = StrSql & "'" & Month(Fechadesde) & Year(Fechadesde) & "',"
                    StrSql = StrSql & "'" & Month(Fechahasta) & Year(Fechahasta) & "',"
                    StrSql = StrSql & "0," 'Renta imponible
                    StrSql = StrSql & "'" & Mid(arregloEstruc(11), 1, 50) & "'," 'Regimen Previsional
                    StrSql = StrSql & arregloEstruc(12) & "," 'Tipo Trabajador
                    StrSql = StrSql & "0," 'Dias Trabajados
                    StrSql = StrSql & "0," 'Codigo Movimiento de personal
                    StrSql = StrSql & "''," 'Fecha Desde para el movimiento
                    StrSql = StrSql & "''," 'Fecha Hasta para el movimiento
                    StrSql = StrSql & "' '," 'Tramo Asignacion Familiar
                    StrSql = StrSql & "0," 'Numero de cargas simples
                    StrSql = StrSql & "0," 'Numero de cargas Maternales
                    StrSql = StrSql & "0," 'Numero de cargas Invalidas
                    StrSql = StrSql & "0," 'ASignacion Familiar
                    StrSql = StrSql & "0," 'ASignacion Familiar Retroactiva
                    StrSql = StrSql & "0," 'Renta Carga Familiares
                    StrSql = StrSql & arregloEstruc(24) & "," 'Codigo AFP
                    StrSql = StrSql & "0," 'Cotizacion Obligatoria AFP
                    StrSql = StrSql & "0," 'Cuenta de ahorro Voluntario
                    StrSql = StrSql & "0," 'Renta Imponible sust AFP
                    StrSql = StrSql & "0.00," 'Tasa Pactada
                    StrSql = StrSql & "0," 'Aporte Indem
                    StrSql = StrSql & "00," 'Numero de Periodos REVISAR!!!!!!!!!!!!!!!!!!!!!
                    StrSql = StrSql & "Null" & ","  'REVISAR CAMBIAR
                    StrSql = StrSql & "Null" & ","  'REVISAR CAMBIAR
                    StrSql = StrSql & "Null" & ","  'Puesto de trabajo Pesado
                    StrSql = StrSql & "00.00," 'Porcentaje Cotizacion Trabajo Pesado
                    StrSql = StrSql & "0," 'Cotizacion Trabajo Pesado
                    StrSql = StrSql & "0" & "," 'RUT Empresa REVISAR!!!!!!!!!!!!!!!!!!!!!
                    StrSql = StrSql & "0" & "," 'DV Empresa REVISAR!!!!!!!!!!!!!!!!!!!!!
                    StrSql = StrSql & arregloEstruc(38) & "," 'Codigo Caja Regimen
                    StrSql = StrSql & "0," 'Tasa cotizacion EX cajas de Regimen
                    StrSql = StrSql & "0," 'Cotizacion Obligatoria INP
                    StrSql = StrSql & "0," 'Renta Imponible Desahucio
                    StrSql = StrSql & "0," 'COdigo ex caja Regimen
                    StrSql = StrSql & "0," 'Tasa Cotizacion Desahucio
                    StrSql = StrSql & "0," 'Cotizacion Desahucio
                    StrSql = StrSql & "0,"    'Otros Aportes INP
                    StrSql = StrSql & "0," 'Cotizacion Fonasa
                    StrSql = StrSql & "0," 'Cotizacion Accidente de Trabajo
                    StrSql = StrSql & "0," 'Bonificacion ley 15.386
                    StrSql = StrSql & "0," 'Descuento por Cargas Familiares
                    StrSql = StrSql & "0," 'Total a Pagar al INP
                    StrSql = StrSql & "0," 'Codigo Institucion de Salud
                    StrSql = StrSql & "0," 'Moneda del plan pactado con Isapre
                    StrSql = StrSql & "0," 'Cotizacion Obligatoria ISAPRE
                    StrSql = StrSql & "0," 'Bonificacion ley 18.566
                    StrSql = StrSql & "0," 'Cotizacion adicional Voluntaria
                    StrSql = StrSql & "0," 'Otros ISAPRE
                    StrSql = StrSql & "0," 'Cotizacion Pactada
                    StrSql = StrSql & "0," 'Total a pagar Isapre
                    StrSql = StrSql & "0,"  'FUN
                    StrSql = StrSql & "0," 'Codigo CCAF
                    StrSql = StrSql & "0," 'Creditos Personales CCAF
                    StrSql = StrSql & "0," 'Descuento Dental
                    StrSql = StrSql & "0," 'Descuento por Leasing
                    StrSql = StrSql & "0," 'Descuentos por Seguro de Vida
                    StrSql = StrSql & "0," 'Otros Descuentos CCAF
                    StrSql = StrSql & "0," 'Cotiz CCAF de no Afiliados ISAPRE
                    StrSql = StrSql & "0," 'Descuentos Cargas Familiares CCAF
                    StrSql = StrSql & "0," 'Codigo Mutual
                    StrSql = StrSql & "0," 'Cotizacion Accidente del trabajo
                    StrSql = StrSql & "0," 'Sucursal Para Pago Mutual
                    StrSql = StrSql & arregloEstruc(71) & "," 'Inst Autor Ahorro Previsonal Vol
                    StrSql = StrSql & arregloEstruc(72) & ","  'Forma de Pago APV
                    StrSql = StrSql & arreglo(73) & "," 'Cotizacion Ahorro Previsonal Vol
                    StrSql = StrSql & arreglo(74) & "," 'Cotizacion Depositos convenidos
                    StrSql = StrSql & "0," 'Renta total Imponible Seguro de Cesantia
                    StrSql = StrSql & "0," 'Aporte Trabajador Seguro de Cesantia
                    StrSql = StrSql & "0," 'Aporte Empleador Seguro de Cesantia
                    StrSql = StrSql & "'" & Mid(arregloEstruc(78), 1, 3) & "',"   'Centro Costo, sucursal, etc
                    StrSql = StrSql & IIf(HuboError, -1, 0)
                    
                    StrSql = StrSql & ")"
                    Flog.writeline
                    Flog.writeline "Insertando : " & StrSql
                    objConn.Execute StrSql, , adExecuteNoRecords
                    Flog.writeline
                    Flog.writeline
                    'Sumo el numero de linea
                    Num_linea = Num_linea + 1
                    Flog.writeline
                    Flog.writeline "SE GRABO LINEA DE GRATIFICACION"
                    Flog.writeline
                                        
                    End If
                    
                    If (SeguroCesantia = True) And (EstaINP = True) Then
                    ' Se arma otra linea adicional
                    StrSql = "INSERT INTO rep_previred (bpronro, ternro, num_linea, Titulo, pliqnro_Desde, pliqnro_hasta, empnro, rut, DV, Apellido, Apellido2, Nombres, sexo, tipo_pago, Periodo_desde,"
                    StrSql = StrSql & "Periodo_hasta, renta_imp, reg_pre, TipTrabajador, DiasTrab, CodmovPer, fechadesde, fechahasta, TramoAsigFam, NumCargasSim, NumCargasMat, "
                    StrSql = StrSql & "NumCargasInv, AsigFam, AsigFamRetro, ReintCarFam, CodAFP, CotizObligAFP, CAVoluntaAFP, RenImpSustAFP, TasaPact, AportIndem, NumPeriodos, "
                    StrSql = StrSql & "PeriDesdeAFP, PeriHastaAFP, PuesTrabPesado, PorcCotizTrabPesa, CotizTrabPesa, RutPag, DVPag, CodCaReg, TasaCotCajPrev, CotizObligINP, "
                    StrSql = StrSql & "RentImpoDesah, CodCaRegDesah, TasaCotDesah, CotizDesah, OtroINP, CotizFonasa, CotizAccTrab, BonLeyInp, DescCargFam, TotPagINP, "
                    StrSql = StrSql & "CodInstSal, MonPlanIsapre, CotizObligIsapre, BonifLeyIsapre, CotizAdicVolun, OtrosIsapre, CotizPact, TotPagIsapre, NumFun, CodCCAF, CredPerCCAF,"
                    StrSql = StrSql & "DescDentCCAF, DescLeasCCAF, DescVidaCCAF, OtrosDesCCAF, CotCCAFnoIsapre, DesCarFamCCAF, CodMut, CotizAccTrabMut, SucPagMut "
                    StrSql = StrSql & ",InstAutAPV, ForPagAPV, CotizAPV, CotizDepConv, RentTotImp, AporTrabSeg, AporEmpSeg, CentroCosto, auxdeci"
                    
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
                    StrSql = StrSql & TipoPago & ","
                    StrSql = StrSql & "'" & Month(Fechadesde) & Year(Fechadesde) & "',"
                    StrSql = StrSql & "'0',"
                    StrSql = StrSql & arreglo(10) & ","    'Renta imponible
                    StrSql = StrSql & "'AFP'," 'Regimen Previsional
                    StrSql = StrSql & arregloEstruc(12) & "," 'Tipo Trabajador
                    StrSql = StrSql & "0," 'Dias Trabajados
                    StrSql = StrSql & "0," 'Codigo Movimiento de personal
                    StrSql = StrSql & "'          '," 'Fecha Desde para el movimiento
                    StrSql = StrSql & "'          '," 'Fecha Hasta para el movimiento
                    StrSql = StrSql & "' '," 'Tramo Asignacion Familiar
                    StrSql = StrSql & "0," 'Numero de cargas simples
                    StrSql = StrSql & "0," 'Numero de cargas Maternales
                    StrSql = StrSql & "0," 'Numero de cargas Invalidas
                    StrSql = StrSql & "0," 'ASignacion Familiar
                    StrSql = StrSql & "0," 'ASignacion Familiar Retroactiva
                    StrSql = StrSql & "0," 'Renta Carga Familiares
                    StrSql = StrSql & arregloEstruc(24) & "," 'Codigo AFP
                    StrSql = StrSql & "0," 'Cotizacion Obligatoria AFP
                    StrSql = StrSql & "0," 'Cuenta de ahorro Voluntario
                    StrSql = StrSql & "0," 'Renta Imponible sust AFP
                    StrSql = StrSql & "0.00," 'Tasa Pactada
                    StrSql = StrSql & "0," 'Aporte Indem
                    StrSql = StrSql & "00," 'Numero de Periodos REVISAR!!!!!!!!!!!!!!!!!!!!!
                    StrSql = StrSql & "Null" & ","  'REVISAR CAMBIAR
                    StrSql = StrSql & "Null" & ","  'REVISAR CAMBIAR
                    StrSql = StrSql & "Null" & ","  'Puesto de trabajo Pesado
                    StrSql = StrSql & "00.00," 'Porcentaje Cotizacion Trabajo Pesado
                    StrSql = StrSql & "0," 'Cotizacion Trabajo Pesado
                    StrSql = StrSql & "Null" & "," 'RUT Empresa REVISAR!!!!!!!!!!!!!!!!!!!!!
                    StrSql = StrSql & "Null" & "," 'DV Empresa REVISAR!!!!!!!!!!!!!!!!!!!!!
                    StrSql = StrSql & "0," 'Codigo Caja Regimen
                    StrSql = StrSql & "0," 'Tasa cotizacion EX cajas de Regimen
                    StrSql = StrSql & "0," 'Cotizacion Obligatoria INP
                    StrSql = StrSql & "0," 'Renta Imponible Desahucio
                    StrSql = StrSql & "0," 'COdigo ex caja Regimen
                    StrSql = StrSql & "0," 'Tasa Cotizacion Desahucio
                    StrSql = StrSql & "0," 'Cotizacion Desahucio
                    StrSql = StrSql & "0,"    'Otros Aportes INP
                    StrSql = StrSql & "0," 'Cotizacion Fonasa
                    StrSql = StrSql & "0," 'Cotizacion Accidente de Trabajo
                    StrSql = StrSql & "0," 'Bonificacion ley 15.386
                    StrSql = StrSql & "0," 'Descuento por Cargas Familiares
                    StrSql = StrSql & "0," 'Total a Pagar al INP
                    StrSql = StrSql & "0," 'Codigo Institucion de Salud
                    StrSql = StrSql & "0," 'Moneda del plan pactado con Isapre
                    StrSql = StrSql & "0," 'Cotizacion Obligatoria ISAPRE
                    StrSql = StrSql & "0," 'Bonificacion ley 18.566
                    StrSql = StrSql & "0," 'Cotizacion adicional Voluntaria
                    StrSql = StrSql & "0," 'Otros ISAPRE
                    StrSql = StrSql & "0," 'Cotizacion Pactada
                    StrSql = StrSql & "0," 'Total a pagar Isapre
                    StrSql = StrSql & "0,"  'FUN
                    StrSql = StrSql & "0," 'Codigo CCAF
                    StrSql = StrSql & "0," 'Creditos Personales CCAF
                    StrSql = StrSql & "0," 'Descuento Dental
                    StrSql = StrSql & "0," 'Descuento por Leasing
                    StrSql = StrSql & "0," 'Descuentos por Seguro de Vida
                    StrSql = StrSql & "0," 'Otros Descuentos CCAF
                    StrSql = StrSql & "0," 'Cotiz CCAF de no Afiliados ISAPRE
                    StrSql = StrSql & "0," 'Descuentos Cargas Familiares CCAF
                    StrSql = StrSql & "0," 'Codigo Mutual
                    StrSql = StrSql & "0," 'Cotizacion Accidente del trabajo
                    StrSql = StrSql & "0," 'Sucursal Para Pago Mutual
                    StrSql = StrSql & "0," 'Inst Autor Ahorro Previsonal Vol
                    StrSql = StrSql & "0,"  'Forma de Pago APV
                    StrSql = StrSql & "0," 'Cotizacion Ahorro Previsonal Vol
                    StrSql = StrSql & "0," 'Cotizacion Depositos convenidos
                    StrSql = StrSql & arreglo(75) & "," 'Renta total Imponible Seguro de Cesantia
                    StrSql = StrSql & arreglo(76) & "," 'Aporte Trabajador Seguro de Cesantia
                    StrSql = StrSql & arreglo(77) & "," 'Aporte Empleador Seguro de Cesantia
                    StrSql = StrSql & "'" & Mid(arregloEstruc(78), 1, 3) & "',"   'Centro Costo, sucursal, etc
                    StrSql = StrSql & IIf(HuboError, -1, 0)
                    
                    StrSql = StrSql & ")"
                    Flog.writeline
                    Flog.writeline "Insertando : " & StrSql
                    objConn.Execute StrSql, , adExecuteNoRecords
                    Flog.writeline
                    Flog.writeline
                    'Sumo el numero de linea
                    Num_linea = Num_linea + 1
                    Flog.writeline
                    Flog.writeline "SE GRABO LINEA ADICIONAL INP CON SEGURO CESANTIA"
                    Flog.writeline
                    
                    End If
                    
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


